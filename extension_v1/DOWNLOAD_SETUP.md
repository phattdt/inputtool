# Extension Download Log - Setup Instructions

## ✅ Đã cập nhật popup.html

File `popup.html` đã có nút **"📥 Download Performance Log"**

## Cần thêm handler vào popup.js

Mở file `extension_v1/popup.js` và thêm đoạn code này vào **cuối file** (sau dòng `});` cuối cùng):

```javascript
// Download log button handler
const downloadBtn = document.getElementById("downloadLog");
if (downloadBtn) {
  downloadBtn.addEventListener("click", () => {
    if (!log.value || log.value.trim() === "") {
      alert("⚠️ No log data to download. Run the extension first!");
      return;
    }

    // Build txt content
    let content = '='.repeat(80) + '\n';
    content += '  📊 EXTENSION PERFORMANCE LOG\n';
    content += '='.repeat(80) + '\n\n';
    content += `Generated: ${new Date().toLocaleString()}\n\n`;
    content += '-'.repeat(80) + '\n';
    content += '  Processing Log:\n';
    content += '-'.repeat(80) + '\n\n';
    content += log.value;
    content += '\n\n' + '='.repeat(80) + '\n';

    // Create and download file
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `extension_performance_log_${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
    a.click();
    URL.revokeObjectURL(url);

    println("📥 Performance log downloaded!");
  });
}
```

## Hoặc copy từ file

Copy toàn bộ nội dung của `download_handler.js` và paste vào cuối `popup.js`

## How to Use

1. Mở extension popup
2. Select PDF file
3. Click "明細書を読み込んで補助"
4. Sau khi xử lý xong → Click **"📥 Download Performance Log"**
5. File `.txt` sẽ được tải về Downloads folder

## Expected Output TXT

```
================================================================================
  📊 EXTENSION PERFORMANCE LOG
================================================================================

Generated: 1/26/2026, 6:00:00 PM

--------------------------------------------------------------------------------
  Processing Log:
--------------------------------------------------------------------------------

1) Extracting PDF text locally…
✅ Extracted 25432 chars in 487.56ms
2) Parsing text into payload JSON…
✅ Parsed parts: 1
✅ Parsed details keys: 1
   Parse time: 256.78ms
3) Injecting fill logic into the current tab…
4) Sending payload to the page (localStorage + call fill)…
✅ Done. Check the Demo-Tool page.
📊 TOTAL TIME: 931.80ms

================================================================================
```

## Nếu không muốn sửa code

Có thể dùng script universal `performance_logger.js` đã tạo trước đó - inject vào page rồi click bookmarklet/extension bình thường.
