document.addEventListener('DOMContentLoaded', () => {
    const timeDisplay = document.getElementById('current-time');
    const employeeIdInput = document.getElementById('employee-id');
    const clockInBtn = document.getElementById('clock-in-btn');
    const clockOutBtn = document.getElementById('clock-out-btn');
    const statusMessage = document.getElementById('status-message');
    const gasUrlInput = document.getElementById('gas-url');
    const saveSettingsBtn = document.getElementById('save-settings');

    // 設定のロード
    const savedGasUrl = localStorage.getItem('attendance_gas_url');
    const savedEmployeeId = localStorage.getItem('attendance_employee_id');

    if (savedGasUrl) gasUrlInput.value = savedGasUrl;
    if (savedEmployeeId) employeeIdInput.value = savedEmployeeId;

    // 時計の更新
    function updateTime() {
        const now = new Date();
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        timeDisplay.textContent = `${hours}:${minutes}`;
    }
    setInterval(updateTime, 1000);
    updateTime();

    // 設定保存
    saveSettingsBtn.addEventListener('click', () => {
        const url = gasUrlInput.value.trim();
        if (url) {
            localStorage.setItem('attendance_gas_url', url);
            showMessage('設定を保存しました', 'success');
        }
    });

    // 社員コード保存（入力が変わるたびに保存）
    employeeIdInput.addEventListener('change', () => {
        localStorage.setItem('attendance_employee_id', employeeIdInput.value.trim());
    });

    // 打刻処理
    async function handleAttendance(type) {
        const employeeId = employeeIdInput.value.trim();
        const gasUrl = localStorage.getItem('attendance_gas_url');

        if (!employeeId) {
            showMessage('社員コードを入力してください', 'error');
            return;
        }

        if (!gasUrl) {
            showMessage('設定からGASアプリのURLを設定してください', 'error');
            return;
        }

        setLoading(true);

        try {
            // GASへの送信データ
            const data = {
                action: type, // 'in' or 'out'
                employeeId: employeeId,
                timestamp: new Date().toISOString()
            };

            // no-corsモードで送信
            await fetch(gasUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            });

            const actionText = type === 'in' ? '出勤' : '退勤';
            showMessage(`${actionText}を記録しました！`, 'success');

        } catch (error) {
            console.error('Error:', error);
            showMessage('送信に失敗しました。ネット接続を確認してください。', 'error');
        } finally {
            setLoading(false);
        }
    }

    function showMessage(msg, type) {
        statusMessage.textContent = msg;
        statusMessage.className = `status-message ${type}`;
        setTimeout(() => {
            statusMessage.textContent = '';
            statusMessage.className = 'status-message';
        }, 5000);
    }

    function setLoading(isLoading) {
        clockInBtn.disabled = isLoading;
        clockOutBtn.disabled = isLoading;
        if (isLoading) {
            statusMessage.textContent = '送信中...';
            statusMessage.className = 'status-message';
        }
    }

    clockInBtn.addEventListener('click', () => handleAttendance('in'));
    clockOutBtn.addEventListener('click', () => handleAttendance('out'));
});
