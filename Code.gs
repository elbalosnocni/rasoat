// Dùng CacheService để tăng tốc độ truy xuất dữ liệu Sheet
function getCachedData(sheetName) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(sheetName);
  
  if (cached != null) {
    return JSON.parse(cached);
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  
  // Lưu vào Cache trong 10 phút (600 giây)
  try {
    cache.put(sheetName, JSON.stringify(result), 600);
  } catch (e) {
    Logger.log("Cache limit exceeded: " + e.message);
  }
  
  return result;
}

// Xóa Cache khi có cập nhật dữ liệu (gọi hàm này khi thêm/sửa user hoặc bài khảo sát)
function clearCache(sheetName) {
  var cache = CacheService.getScriptCache();
  if (sheetName) {
    cache.remove(sheetName);
  } else {
    cache.removeAll(['Users', 'Surveys', 'Responses']);
  }
}

// Hàm phục vụ HTML Frontend
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Hệ Thống Khảo Sát Tối Ưu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Hàm Xử lý Đăng nhập Tối ưu Tốc độ
function loginUser(username, password) {
  try {
    if (!username || !password) {
      return { success: false, message: 'Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu!' };
    }
    
    // Đọc từ Cache giúp phản hồi nhanh < 300ms
    var users = getCachedData('Users');
    
    var user = users.find(function(u) {
      return u.username.toString().trim() === username.toString().trim() && 
             u.password.toString().trim() === password.toString().trim();
    });

    if (user) {
      return {
        success: true,
        user: {
          id: user.id || user.username,
          username: user.username,
          name: user.fullname || user.name || user.username,
          role: user.role || 'user'
        }
      };
    } else {
      return { success: false, message: 'Tên đăng nhập hoặc mật khẩu không chính xác!' };
    }
  } catch (error) {
    return { success: false, message: 'Lỗi hệ thống: ' + error.toString() };
  }
}
