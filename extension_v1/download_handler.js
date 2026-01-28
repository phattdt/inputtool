// Add this code at the END of popup.js

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
