/**
 * ╔══════════════════════════════════════════════════════════════╗
 * ║  HR-NEXUS Cloud Trigger v1.0                                ║
 * ║  Kích hoạt Cloud Run Job từ Google Sheets                   ║
 * ║  Thay thế các nút GAS cũ bằng Python Engine trên mây       ║
 * ╚══════════════════════════════════════════════════════════════╝
 *
 * CÁCH SỬ DỤNG:
 * 1. Mở Extensions → Apps Script trên Google Sheet
 * 2. Paste đoạn code này vào cuối file Code.gs hiện tại
 * 3. Trong hàm onOpen() hiện tại, thêm: addCloudMenuItem(menu);
 * 4. Cập nhật appsscript.json (xem hướng dẫn bên dưới)
 * 5. Lưu → Chạy → Authorize
 */

// ─── CẤU HÌNH CLOUD ────────────────────────────────────────────
const CLOUD_CONFIG = {
  PROJECT_ID: '<YOUR_GCP_PROJECT_ID_HERE>',
  LOCATION: '<YOUR_GCP_REGION_HERE>',
  JOB_NAME: '<YOUR_CLOUD_RUN_JOB_NAME_HERE>',
  POLL_INTERVAL_MS: 10000,  // Poll mỗi 10 giây
  MAX_POLL_ATTEMPTS: 36,     // Tối đa 6 phút (36 x 10s)
};


// ─── MENU INTEGRATION ──────────────────────────────────────────

/**
 * GỌI HÀM NÀY TRONG onOpen() HIỆN TẠI CỦA ANH
 * 
 * Ví dụ trong onOpen() hiện tại:
 *   function onOpen() {
 *     var ui = SpreadsheetApp.getUi();
 *     var menu = ui.createMenu('L&D Vận hành v3.4.2');
 *     // ... các menu item cũ ...
 *     addCloudMenuItem(menu);  // ← THÊM DÒNG NÀY
 *     menu.addToUi();
 *   }
 */
function addCloudMenuItem(menu) {
  menu.addSeparator()
      .addItem('🚀 Chạy báo cáo Cloud (Python Engine)', 'runCloudReport')
      .addItem('📊 Kiểm tra trạng thái Cloud Job', 'checkCloudJobStatus');
}


// ─── NÚT CHÍNH: CHẠY BÁO CÁO CLOUD ───────────────────────────

/**
 * Kích hoạt Cloud Run Job - Thay thế tất cả nút đồng bộ cũ
 * Menu: "🚀 Chạy báo cáo Cloud (Python Engine)"
 */
function runCloudReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ── Step 1: Xác nhận với Admin ──
  var confirm = ui.alert(
    '🚀 Chạy Báo Cáo Cloud',
    'Hệ thống Python Engine trên mây sẽ:\n\n' +
    '  ✦ Kéo toàn bộ dữ liệu từ Google Sheet\n' +
    '  ✦ Đồng bộ fact table (Training Records)\n' +
    '  ✦ Tính toán 10 sheet phân tích báo cáo\n' +
    '  ✦ Ghi kết quả trở lại Google Sheet\n\n' +
    '⏱️ Thời gian ước tính: 1-2 phút.\n\n' +
    'Tiếp tục?',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (confirm !== ui.Button.OK) return;
  
  // ── Step 2: Gửi lệnh lên Cloud ──
  ss.toast('⏳ Đang gửi lệnh lên Cloud...', '☁️ Cloud Engine', 5);
  
  try {
    var operationName = executeCloudRunJob_();
    
    ss.toast('✅ Đã kích hoạt! Đang chờ kết quả...', '☁️ Cloud Engine', 10);
    
    // ── Step 3: Poll trạng thái cho đến khi xong ──
    var result = pollExecution_(operationName);
    
    // ── Step 4: Hiển thị kết quả ──
    if (result.success) {
      ss.toast('🎉 Hoàn tất!', '☁️ Cloud Engine', 10);
      ui.alert(
        '✅ Báo Cáo Đã Cập Nhật!',
        '🎉 Python Engine đã xử lý thành công!\n\n' +
        '⏱️ Thời gian xử lý: ' + result.duration + '\n\n' +
        '📊 Hãy kiểm tra các sheet báo cáo nhé.\n\n' +
        '💡 Gợi ý: Nhấn Ctrl+Shift+E để mở lại sheet.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '⚠️ Cloud Job Gặp Lỗi',
        'Chi tiết lỗi:\n' + result.error + '\n\n' +
        '💡 Gợi ý:\n' +
        '• Thử bấm lại nút "🚀 Chạy báo cáo Cloud"\n' +
        '• Nếu vẫn lỗi, liên hệ IT và gửi đoạn lỗi trên.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (e) {
    ss.toast('❌ Lỗi: ' + e.message, '☁️ Cloud Engine', 10);
    ui.alert(
      '❌ Lỗi Kết Nối',
      'Không thể kết nối Cloud Engine:\n\n' + e.message + '\n\n' +
      '💡 Kiểm tra:\n' +
      '• Kết nối mạng Internet\n' +
      '• Quyền truy cập Google Cloud',
      ui.ButtonSet.OK
    );
  }
}


// ─── KIỂM TRA TRẠNG THÁI ──────────────────────────────────────

/**
 * Kiểm tra trạng thái Cloud Job (để debug)
 * Menu: "📊 Kiểm tra trạng thái Cloud Job"
 */
function checkCloudJobStatus() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var url = 'https://run.googleapis.com/v2/projects/' +
              CLOUD_CONFIG.PROJECT_ID + '/locations/' +
              CLOUD_CONFIG.LOCATION + '/jobs/' +
              CLOUD_CONFIG.JOB_NAME;
              
    var token = ScriptApp.getOAuthToken();
    
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    
    if (response.getResponseCode() !== 200) {
      ui.alert('⚠️ Lỗi', 'Cloud Job chưa được tạo hoặc không truy cập được.\n' +
               'Code: ' + response.getResponseCode(), ui.ButtonSet.OK);
      return;
    }
    
    var data = JSON.parse(response.getContentText());
    
    var statusMsg = '☁️ TRẠNG THÁI CLOUD JOB\n';
    statusMsg += '━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
    statusMsg += '• Job: ' + CLOUD_CONFIG.JOB_NAME + '\n';
    statusMsg += '• Khu vực: ' + CLOUD_CONFIG.LOCATION + '\n';
    statusMsg += '• Project: ' + CLOUD_CONFIG.PROJECT_ID + '\n\n';
    
    if (data.latestCreatedExecution) {
      var exec = data.latestCreatedExecution;
      statusMsg += '📌 LẦN CHẠY GẦN NHẤT:\n';
      statusMsg += '• Thời gian: ' + (exec.createTime || 'N/A') + '\n';
      statusMsg += '• Trạng thái: ' + (exec.completionTime ? '✅ Hoàn tất' : '⏳ Đang chạy') + '\n';
    } else {
      statusMsg += '📌 Chưa có lần chạy nào.\n';
    }
    
    statusMsg += '\n• Tổng số lần chạy: ' + (data.executionCount || 0);
    
    ui.alert('☁️ Cloud Job Info', statusMsg, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('❌ Lỗi', 'Không thể kiểm tra: ' + e.message, ui.ButtonSet.OK);
  }
}


// ─── HELPER: GỌI CLOUD RUN API ────────────────────────────────

/**
 * Gọi Cloud Run Jobs API để execute job
 * @returns {string} Operation name (dùng để poll trạng thái)
 * @private
 */
function executeCloudRunJob_() {
  var url = 'https://run.googleapis.com/v2/projects/' +
            CLOUD_CONFIG.PROJECT_ID + '/locations/' +
            CLOUD_CONFIG.LOCATION + '/jobs/' +
            CLOUD_CONFIG.JOB_NAME + ':run';
  
  var token = ScriptApp.getOAuthToken();
  
  var options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({}),
    muteHttpExceptions: true,
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  
  if (code !== 200) {
    var body = response.getContentText();
    throw new Error('Cloud API lỗi ' + code + ': ' + body.substring(0, 200));
  }
  
  var data = JSON.parse(response.getContentText());
  
  // Cloud Run Jobs API trả về Operation name
  var opName = data.name || '';
  if (!opName) {
    throw new Error('Không nhận được Operation ID từ Cloud API');
  }
  
  return opName;
}


/**
 * Poll trạng thái execution cho đến khi hoàn tất
 * @param {string} operationName - Cloud Run operation name
 * @returns {Object} {success: boolean, duration: string, error: string}
 * @private
 */
function pollExecution_(operationName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var token = ScriptApp.getOAuthToken();
  
  for (var i = 0; i < CLOUD_CONFIG.MAX_POLL_ATTEMPTS; i++) {
    // Đợi 10 giây giữa mỗi lần kiểm tra
    Utilities.sleep(CLOUD_CONFIG.POLL_INTERVAL_MS);
    
    var url = 'https://run.googleapis.com/v2/' + operationName;
    
    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true,
      });
      
      var data = JSON.parse(response.getContentText());
      
      // Cập nhật tiến trình trên thanh toast
      var elapsed = Math.round((i + 1) * CLOUD_CONFIG.POLL_INTERVAL_MS / 1000);
      ss.toast('⏳ Đang xử lý... (' + elapsed + 's)', '☁️ Cloud Engine', 12);
      
      // Kiểm tra xem Operation đã hoàn tất chưa
      if (data.done) {
        if (data.error) {
          return { success: false, duration: duration, error: data.error.message };
        }
        
        var exec = data.response;
        var startTime = new Date(exec.createTime || Date.now());
        var endTime = new Date(exec.completionTime || Date.now());
        var duration = Math.round((endTime - startTime) / 1000) + ' giây';
        
        var conditions = exec.conditions || [];
        for (var j = 0; j < conditions.length; j++) {
          if (conditions[j].type === 'Completed' && conditions[j].state === 'CONDITION_FAILED') {
            return { success: false, duration: duration, error: conditions[j].message };
          }
        }
        return { success: true, duration: duration, error: '' };
      }
      
      // Nếu có lỗi rõ ràng
      if (data.error) {
        return {
          success: false,
          duration: elapsed + ' giây',
          error: data.error.message || JSON.stringify(data.error),
        };
      }
      
    } catch (e) {
      // Lỗi mạng tạm thời, tiếp tục polling
      ss.toast('⏳ Đang chờ phản hồi... (lần ' + (i + 1) + ')', '☁️ Cloud Engine', 12);
    }
  }
  
  // Hết thời gian polling
  return {
    success: false,
    duration: 'N/A',
    error: 'Timeout - Job chạy quá lâu (>' + 
           Math.round(CLOUD_CONFIG.MAX_POLL_ATTEMPTS * CLOUD_CONFIG.POLL_INTERVAL_MS / 60000) + 
           ' phút). Kiểm tra tại: console.cloud.google.com',
  };
}


// ─── HƯỚNG DẪN CÀI ĐẶT ────────────────────────────────────────
//
// BƯỚC 1: Mở Extensions → Apps Script trên Google Sheet
//
// BƯỚC 2: Paste đoạn code này vào CUỐI file Code.gs hiện tại
//
// BƯỚC 3: Tìm hàm onOpen() hiện tại, thêm dòng sau VÀO CUỐI
//         (TRƯỚC dòng menu.addToUi()):
//
//     addCloudMenuItem(menu);
//
// BƯỚC 4: Click vào file appsscript.json (ở sidebar trái)
//         Nếu không thấy, vào Settings → check "Show appsscript.json"
//         Thêm scope này vào mảng "oauthScopes":
//
//     "https://www.googleapis.com/auth/cloud-platform"
//
// BƯỚC 5: Lưu (Ctrl+S) → Chạy hàm onOpen() 1 lần → Authorize
//
// XONG! Reload Google Sheet, menu mới sẽ xuất hiện.
// ────────────────────────────────────────────────────────────────
