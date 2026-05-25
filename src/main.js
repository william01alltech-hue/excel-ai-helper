// 儲存全域狀態
let officeInitialized = false;
let ollamaHost = 'http://localhost:11434'; // 直接連線至本地 Ollama 服務
let selectedModel = '';
let chatHistory = [];
let lastParsedData = null; // 儲存準備填入 Excel 的二維陣列數據

// 初始化 Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    officeInitialized = true;
    updateStatus('Excel 增益集載入成功，準備就緒。');
    
    // 初始化完成後，嘗試連線本地 Ollama
    initOllama();
  } else {
    updateStatus('目前不在 Excel 中運行，部分 Excel 功能可能受限。');
    initOllama();
  }
});

// UI 元素
const connectionBadge = document.getElementById('connection-badge');
const hostInput = document.getElementById('ollama-host');
const modelSelect = document.getElementById('model-select');
const currentModelName = document.getElementById('current-model-name');
const btnRefreshModels = document.getElementById('btn-refresh-models');
const chatMessages = document.getElementById('chat-messages');
const userInput = document.getElementById('user-input');
const btnSend = document.getElementById('btn-send');
const btnClearChat = document.getElementById('btn-clear-chat');
const btnAnalyzeFormat = document.getElementById('btn-analyze-format');
const btnFillExcel = document.getElementById('btn-fill-excel');
const statusText = document.getElementById('status-text');

// 綁定事件監聽器
btnSend.addEventListener('click', handleSendMessage);
userInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    handleSendMessage();
  }
});
btnClearChat.addEventListener('click', clearChat);
btnRefreshModels.addEventListener('click', fetchModels);
btnAnalyzeFormat.addEventListener('click', analyzeExcelSelection);
btnFillExcel.addEventListener('click', fillDataToExcel);

// 更新狀態文字
function updateStatus(msg) {
  statusText.textContent = msg;
  console.log(`[Status] ${msg}`);
}

// 初始化 Ollama
async function initOllama() {
  // 網址變更處理
  hostInput.addEventListener('change', () => {
    const val = hostInput.value.trim();
    ollamaHost = val || 'http://localhost:11434';
    fetchModels();
  });
  
  // 取得模型清單
  await fetchModels();
}

// 取得 Ollama 模型清單
async function fetchModels() {
  updateStatus('正在連線至本地 Ollama 服務...');
  connectionBadge.className = 'badge badge-disconnected';
  connectionBadge.textContent = '連線中...';
  
  try {
    const response = await fetch(`${ollamaHost}/api/tags`, {
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    });
    
    if (!response.ok) {
      throw new Error(`HTTP 錯誤: ${response.status}`);
    }
    
    const data = await response.json();
    modelSelect.innerHTML = '';
    
    if (data.models && data.models.length > 0) {
      data.models.forEach(model => {
        const option = document.createElement('option');
        option.value = model.name;
        option.textContent = model.name;
        modelSelect.appendChild(option);
      });
      
      // 預設選取第一個模型
      selectedModel = data.models[0].name;
      currentModelName.textContent = selectedModel;
      
      connectionBadge.className = 'badge badge-connected';
      connectionBadge.textContent = '已連線';
      updateStatus('Ollama 連線成功。');
      
      // 監聽選擇變更
      modelSelect.addEventListener('change', () => {
        selectedModel = modelSelect.value;
        currentModelName.textContent = selectedModel;
      });
    } else {
      const option = document.createElement('option');
      option.value = '';
      option.textContent = '無可用模型 (請先 ollama pull)';
      modelSelect.appendChild(option);
      
      connectionBadge.className = 'badge badge-disconnected';
      connectionBadge.textContent = '無模型';
      updateStatus('已連線 Ollama，但未找到已下載的 AI 模型。');
    }
  } catch (error) {
    console.error('Ollama 連線失敗:', error);
    modelSelect.innerHTML = '<option value="">連線失敗，請檢查服務是否啟動</option>';
    connectionBadge.className = 'badge badge-disconnected';
    connectionBadge.textContent = '未連線';
    updateStatus('無法連線至 Ollama。請確認終端機是否已執行 ollama serve。');
  }
}

// 新增訊息到對話框
function appendMessage(role, content) {
  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${role}-message`;
  
  const contentDiv = document.createElement('div');
  contentDiv.className = 'message-content';
  
  // 簡易將換行與程式碼區塊做格式化
  contentDiv.innerHTML = formatMarkdown(content);
  
  messageDiv.appendChild(contentDiv);
  chatMessages.appendChild(messageDiv);
  
  // 滾動到底部
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

// 簡易 Markdown 解析 (處理 pre / code 與換行)
function formatMarkdown(text) {
  // 安全地轉義 HTML
  let escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  
  // 解析 ```json ... ``` 或 ```javascript ... ``` 區塊
  escaped = escaped.replace(/```(?:json|javascript|js|vba|)?\n([\s\S]*?)\n```/g, '<pre><code>$1</code></pre>');
  
  // 解析行內 `code`
  escaped = escaped.replace(/`([^`]+)`/g, '<code>$1</code>');
  
  // 換行
  return escaped.replace(/\n/g, '<br>');
}

// 清空對話
function clearChat() {
  chatMessages.innerHTML = '';
  chatHistory = [];
  lastParsedData = null;
  btnFillExcel.disabled = true;
  updateStatus('對話紀錄已清除。');
}

// 處理傳送對話
async function handleSendMessage() {
  const text = userInput.value.trim();
  if (!text) return;
  
  if (!selectedModel) {
    alert('請先在上方設定並選取一個 Ollama 模型！');
    return;
  }
  
  // 顯示使用者訊息
  appendMessage('user', text);
  userInput.value = '';
  
  // 加入歷史紀錄
  chatHistory.push({ role: 'user', content: text });
  
  // 顯示 AI 載入中動畫
  const loadingDiv = document.createElement('div');
  loadingDiv.className = 'message assistant-message';
  loadingDiv.id = 'ai-loading-bubble';
  loadingDiv.innerHTML = `
    <div class="message-content">
      <div class="loading-dots">
        <span></span><span></span><span></span>
      </div>
    </div>
  `;
  chatMessages.appendChild(loadingDiv);
  chatMessages.scrollTop = chatMessages.scrollHeight;
  
  updateStatus('AI 思考中...');
  
  try {
    const response = await fetch(`${ollamaHost}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: selectedModel,
        messages: chatHistory,
        stream: false
      })
    });
    
    // 移除載入動畫
    const bubble = document.getElementById('ai-loading-bubble');
    if (bubble) bubble.remove();
    
    if (!response.ok) {
      throw new Error(`HTTP 錯誤: ${response.status}`);
    }
    
    const data = await response.json();
    const aiResponse = data.message.content;
    
    // 顯示 AI 回應
    appendMessage('assistant', aiResponse);
    
    // 存入歷史
    chatHistory.push({ role: 'assistant', content: aiResponse });
    
    // 檢查回應中是否有 JSON 陣列資料，若有，則解析並提供「寫入 Excel」功能
    detectAndParseJsonData(aiResponse);
    
    updateStatus('AI 回應完畢。');
  } catch (error) {
    console.error('AI 請求失敗:', error);
    const bubble = document.getElementById('ai-loading-bubble');
    if (bubble) bubble.remove();
    
    appendMessage('system', `連線錯誤：無法從 Ollama 取得回應。請檢查您的終端機是否開啟，或是否有設定 OLLAMA_ORIGINS="*"。`);
    updateStatus('請求失敗。');
  }
}

// 偵測並解析 AI 回傳的 JSON 資料（主要是二維陣列）
function detectAndParseJsonData(text) {
  // 使用正則表達式尋找 Markdown 中的 JSON 區塊
  const jsonRegex = /```json\s*([\s\S]*?)\s*```/;
  const match = text.match(jsonRegex);
  
  let jsonString = '';
  if (match && match[1]) {
    jsonString = match[1].trim();
  } else {
    // 嘗試尋找大括號或中括號包裹的字串
    const arrayRegex = /(\[[\s\S]*?\])/;
    const arrayMatch = text.match(arrayRegex);
    if (arrayMatch) {
      jsonString = arrayMatch[0].trim();
    }
  }
  
  if (jsonString) {
    try {
      const parsed = JSON.parse(jsonString);
      // 驗證是否為二維陣列 (例如: [["欄位1", "欄位2"], ["資料1", "資料2"]])
      if (Array.isArray(parsed) && parsed.length > 0 && Array.isArray(parsed[0])) {
        lastParsedData = parsed;
        btnFillExcel.disabled = false;
        updateStatus('偵測到可填入 Excel 的表格資料！您可以點擊下方綠色按鈕寫入。');
        appendMessage('system', '✨ 系統偵測到表格資料，下方的「將 AI 生成填入 Excel」按鈕已啟用，點擊可直接填入您當前選取的位置！');
      }
    } catch (e) {
      // 解析失敗不報錯，可能只是一般的 JSON 程式碼
      console.log('未偵測到合法的二維數據陣列:', e);
    }
  }
}

// Excel 工具功能：讀取選取區域的格式
async function analyzeExcelSelection() {
  if (!officeInitialized) {
    alert('Excel 增益集尚未與 Excel 軟體完成綁定！請確認您是在 Excel 內打開此側邊欄。');
    return;
  }
  
  updateStatus('正在讀取選取範圍格式...');
  
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "address", "columnCount", "rowCount"]);
      await context.sync();
      
      const values = range.values;
      const address = range.address;
      
      if (!values || values.length === 0 || (values.length === 1 && values[0][0] === "")) {
        appendMessage('system', '您的選取範圍是空的。請先在 Excel 中選取含有假單或工時表標頭的範圍。');
        updateStatus('選取範圍無有效內容。');
        return;
      }
      
      // 將選取的網格格式化為提示詞
      let gridText = '';
      values.forEach((row, rIdx) => {
        gridText += `第 ${rIdx + 1} 行: [${row.map(cell => cell === "" ? "(空白)" : `"${cell}"`).join(', ')}]\n`;
      });
      
      const prompt = `我目前在 Excel 中選取了範圍 ${address}，它的儲存格格式與現有文字如下：
\`\`\`
${gridText}\`\`\`
這是一個假單或工時表的範本格式。請幫我分析這個格式，並產生 1 到 3 筆合理的模擬假單/工時表資料。

**重要規則**：
1. 請以「繁體中文」生成合理的資料（如姓名、事由、時間等）。
2. 除了文字說明外，請務必在回應的最後附上一個 JSON 格式的二維陣列 (Array of Arrays)，格式如下，以便我的程式能直接寫入儲存格中：
\`\`\`json
[
  ["姓名", "日期", "假別", "事由"],
  ["張小明", "2026-05-26", "事假", "處理家庭事務"]
]
\`\`\`
（注意：JSON 內部的每一行必須對應您生成的資料列。如果您選取的範圍包含了表頭，JSON 中可以只包含純資料列，或者包含表頭加資料列。請盡量配合我所選取的格式欄位對齊。）`;

      // 自動填入輸入框並觸發發送
      userInput.value = prompt;
      handleSendMessage();
    });
  } catch (error) {
    console.error('讀取 Excel 失敗:', error);
    updateStatus('讀取選取範圍失敗。');
    appendMessage('system', `讀取工作表失敗: ${error.message}`);
  }
}

// Excel 工具功能：將 AI 生成的二維陣列填入 Excel
async function fillDataToExcel() {
  if (!officeInitialized) {
    alert('Excel 增益集尚未完成初始化！');
    return;
  }
  
  if (!lastParsedData || lastParsedData.length === 0) {
    alert('目前沒有已解析的 AI 生成數據。請先讓 AI 生成資料。');
    return;
  }
  
  updateStatus('正在將資料寫入 Excel...');
  
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      
      // 取得選取範圍的左上角儲存格，並調整為 AI 生成的陣列大小
      const activeCell = range.getCell(0, 0);
      const rowCount = lastParsedData.length;
      const colCount = lastParsedData[0].length;
      
      // 取得對應尺寸的目標範圍 (getResizedRange 的參數是列增量與行增量，因此需減 1)
      const targetRange = activeCell.getResizedRange(rowCount - 1, colCount - 1);
      targetRange.values = lastParsedData;
      
      // 自動調整欄寬
      targetRange.format.autofitColumns();
      
      await context.sync();
      updateStatus('寫入 Excel 成功！');
      appendMessage('system', `成功將 ${rowCount} 行、${colCount} 列的資料填入工作表，位置自 ${range.address.split('!')[1].split(':')[0]} 開始。`);
    });
  } catch (error) {
    console.error('寫入 Excel 失敗:', error);
    updateStatus('寫入失敗。');
    appendMessage('system', `寫入工作表失敗: ${error.message}`);
  }
}
