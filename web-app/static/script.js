document.getElementById('scheduleForm').addEventListener('submit', function(e) {
    const status = document.getElementById('status');
    status.textContent = '正在生成排班表，请稍候...';
});