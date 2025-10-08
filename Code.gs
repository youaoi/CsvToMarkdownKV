/**
 * @fileoverview CsvToMarkdownKV Webアプリケーションのサーバーサイドロジック。
 * GASのdoGetトリガーでUIを配信し、クライアントからのリクエストに応じてファイル変換処理を実行します。
 */

/**
 * HTTP GETリクエストを受信した際にWebページを生成して返します。
 * @param {GoogleAppsScript.Events.DoGet} e イベントオブジェクト。
 * @returns {GoogleAppsScript.HTML.HtmlOutput} ユーザーに表示するHTMLページ。
 */
function doGet(e) {
  const htmlTemplate = HtmlService.createTemplateFromFile('index.html');
  htmlTemplate.css = HtmlService.createHtmlOutputFromFile('stylesheet.html').getContent();
  htmlTemplate.js = HtmlService.createHtmlOutputFromFile('javascript.html').getContent();
  return htmlTemplate.evaluate()
    .setTitle('CsvToMarkdownKV Converter')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Base64エンコードされた複数のファイルデータを受け取り、指定されたフォーマットに一括変換します。
 * @param {Array<Object>} fileDataArray ファイル名とBase64エンコードされた内容を持つオブジェクトの配列。
 * 各オブジェクトは {filename: string, base64Content: string} の形式。
 * @param {string} outputFormat 変換先のフォーマット ('markdown', 'yaml', 'xml')。
 * @returns {Array<Object>|Object} 変換成功時は、変換後のファイル名と内容を含むオブジェクトの配列。
 * 失敗時は、エラーメッセージを含むオブジェクト。
 */
function convertFiles(fileDataArray, outputFormat) {
  if (!fileDataArray || fileDataArray.length === 0 || !outputFormat) {
    return { error: '変換するファイルまたは出力形式が指定されていません。' };
  }

  const convertedResults = [];
  const errorLogs = []; 

  for (const fileData of fileDataArray) {
    try {
      if (!fileData.base64Content) {
        throw new Error("ファイルの内容が空です。");
      }
      const bytes = Utilities.base64Decode(fileData.base64Content);
      if (bytes.length === 0) {
        throw new Error("ファイルが空か、内容を読み取れませんでした。");
      }

      const textContent = decodeBytesToString(bytes);
      const records = parseCsvTsv(textContent);
      
      if (records.length === 0) {
        console.warn(`File "${fileData.filename}" contained no valid data rows after parsing.`);
        continue;
      }

      let fileContent = '';
      let fileExtension = '';
      switch (outputFormat) {
        case 'markdown':
          fileContent = toMarkdownKv(records);
          fileExtension = 'md';
          break;
        case 'yaml':
          fileContent = toYaml(records);
          fileExtension = 'yaml';
          break;
        case 'xml':
          fileContent = toXml(records);
          fileExtension = 'xml';
          break;
        default:
          throw new Error('サポートされていない出力形式です: ' + outputFormat);
      }
      const baseName = fileData.filename.split('.').slice(0, -1).join('.') || fileData.filename;
      const newFilename = `${baseName}.${fileExtension}`;
      
      convertedResults.push({
        filename: newFilename,
        fileContent: fileContent
      });

    } catch (e) {
      const errorMessage = `ファイル「${fileData.filename}」の処理中にエラー: ${e.message}`;
      console.error(errorMessage, e.stack);
      errorLogs.push(errorMessage);
    }
  }

  if (convertedResults.length === 0 && errorLogs.length > 0) {
    return { error: errorLogs.join('\n') };
  }
  
  if (convertedResults.length === 0) {
    return { error: 'どのファイルにも変換可能なデータが含まれていませんでした。' };
  }

  return convertedResults;
}

/**
 * バイト配列を文字列にデコードします。UTF-8を試し、デコードエラーの兆候があればShift_JISを試します。
 * @param {byte[]} bytes ファイルから読み込んだバイト配列。
 * @returns {string} デコードされた文字列。
 * @throws {Error} UTF-8とShift_JISの両方でデコードに失敗した場合。
 */
function decodeBytesToString(bytes) {
  let textContent;
  let decodedAsUtf8;
  
  try {
    decodedAsUtf8 = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    
    // Unicodeの置換文字(U+FFFD)が含まれている場合、デコードエラーと見なす
    if (decodedAsUtf8.includes('\uFFFD')) {
      textContent = Utilities.newBlob(bytes).getDataAsString('Shift_JIS');
    } else {
      textContent = decodedAsUtf8;
    }
  } catch (e) {
    try {
      textContent = Utilities.newBlob(bytes).getDataAsString('Shift_JIS');
    } catch (sjisError) {
      throw new Error(`UTF-8とShift_JISの両方でファイルのデコードに失敗しました。Error: ${sjisError.message}`);
    }
  }
  
  return textContent;
}


/**
 * CSV/TSV形式の文字列を、キーと値を持つオブジェクトの配列に変換します。
 * GAS標準のCSV解析機能 `Utilities.parseCsv()` を利用します。
 * @param {string} text CSV/TSV形式の文字列データ。
 * @returns {Array<Object>} パースされたデータの配列。
 * @throws {Error} CSV解析に失敗した場合。
 */
function parseCsvTsv(text) {
  if (typeof text !== 'string' || !text.trim()) {
    return [];
  }
  let cleanText = text.trim();
  // BOM (Byte Order Mark) を除去
  if (cleanText.charCodeAt(0) === 0xFEFF) {
    cleanText = cleanText.substring(1);
  }
  
  // 1行目から区切り文字を自動判定 (タブ or カンマ)
  const firstLine = cleanText.substring(0, cleanText.indexOf('\n'));
  const delimiter = firstLine.includes('\t') ? '\t' : ',';

  try {
    const data = Utilities.parseCsv(cleanText, delimiter);
    if (!data || data.length < 2) return [];
    
    const headers = data.shift();
    return data
      .map(row => {
        // 空行は無視する
        if (row.length === 0 || row.every(cell => !cell)) return null;
        const record = {};
        headers.forEach((header, index) => {
          const key = header ? header.trim() : `column_${index + 1}`;
          record[key] = (index < row.length && row[index]) ? row[index].trim() : '';
        });
        return record;
      })
      .filter(record => record !== null); // nullになった空行を除外
  } catch(e) {
    throw new Error(`CSV解析に失敗しました。(${e.message})`);
  }
}

/**
 * オブジェクト配列をMarkdownのKey-Value形式の文字列に変換します。
 * @param {Array<Object>} records 変換するデータの配列。
 * @returns {string} Markdown形式の文字列。
 */
function toMarkdownKv(records){
  return records.map(record => {
    return Object.entries(record)
      .map(([key, value]) => `- ${key}: ${value}`)
      .join('\n') + '\n---';
  }).join('\n');
}

/**
 * オブジェクト配列をYAML形式の文字列に変換します。
 * @param {Array<Object>} records 変換するデータの配列。
 * @returns {string} YAML形式の文字列。
 */
function toYaml(records){
  return records.map(record => {
    return '- ' + Object.entries(record)
      .map(([key, value]) => {
        const formattedValue = (typeof value === 'string' && /[:\s]/.test(value)) ? `"${value}"` : value;
        return `  ${key}: ${formattedValue}`;
      })
      .join('\n');
  }).join('\n');
}

/**
 * オブジェクト配列をXML形式の文字列に変換します。
 * @param {Array<Object>} records 変換するデータの配列。
 * @returns {string} XML形式の文字列。
 */
function toXml(records){
  const items = records.map(record => {
    const elements = Object.entries(record).map(([key, value]) => {
      const safeKey = key.replace(/[^a-zA-Z0-9_]/g, '');
      const safeValue = typeof value === 'string' ? value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') : value;
      return `    <${safeKey}>${safeValue}</${safeKey}>`;
    }).join('\n');
    return `  <item>\n${elements}\n  </item>`;
  }).join('\n');
  
  return `<?xml version="1.0" encoding="UTF-8"?>\n<root>\n${items}\n</root>`;
}