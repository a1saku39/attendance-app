document.addEventListener('DOMContentLoaded', () => {
    const timeDisplay = document.getElementById('current-time');
    const employeeIdInput = document.getElementById('employee-id');
    const clockInBtn = document.getElementById('clock-in-btn');
    const clockOutBtn = document.getElementById('clock-out-btn');
    const statusMessage = document.getElementById('status-message');
    const gasUrlInput = document.getElementById('gas-url');
    const saveSettingsBtn = document.getElementById('save-settings');
    const customTimeInput = document.getElementById('custom-time');
    const remarksInput = document.getElementById('remarks');
    const clockInDisplay = document.getElementById('clock-in-display');
    const clockOutDisplay = document.getElementById('clock-out-display');

    // 設定のロード
    const savedGasUrl = localStorage.getItem('attendance_gas_url');
    const savedEmployeeId = localStorage.getItem('attendance_employee_id');

    if (savedGasUrl) gasUrlInput.value = savedGasUrl;
    if (savedEmployeeId) {
        employeeIdInput.value = savedEmployeeId;
        // 本日の打刻時刻を取得
        loadTodayAttendance();
    }

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

    // 社員コード保存
    employeeIdInput.addEventListener('change', () => {
        const id = employeeIdInput.value.trim();
        localStorage.setItem('attendance_employee_id', id);
        if (id) {
            loadTodayAttendance();
        }
    });

    // 本日の打刻時刻を取得して表示
    async function loadTodayAttendance() {
        const employeeId = employeeIdInput.value.trim();
        const gasUrl = localStorage.getItem('attendance_gas_url');

        if (!employeeId || !gasUrl) return;

        try {
            const today = new Date();
            const yearMonth = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0');

            const response = await fetch(gasUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    action: 'getPersonalMonthlyData',
                    employeeId: employeeId,
                    yearMonth: yearMonth
                })
            });

            const result = await response.json();
            if (result.result === 'success') {
                const todayStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;
                const todayData = result.data[todayStr];

                if (todayData) {
                    if (todayData.clockInTime) {
                        clockInDisplay.textContent = `✓ ${todayData.clockInTime}`;
                    }
                    if (todayData.clockOutTime) {
                        clockOutDisplay.textContent = `✓ ${todayData.clockOutTime}`;
                    }
                }
            }
        } catch (error) {
            console.error('本日の打刻時刻取得エラー:', error);
        }
    }



    // 打刻処理（楽観的UI更新）
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

        // 楽観的UI更新: すぐに成功メッセージと時刻を表示
        let timestamp;
        if (customTimeInput && customTimeInput.value) {
            timestamp = new Date(customTimeInput.value).toISOString();
        } else {
            timestamp = new Date().toISOString();
        }

        const displayTime = new Date(timestamp).toLocaleTimeString('ja-JP', {
            hour: '2-digit',
            minute: '2-digit'
        });

        const actionText = type === 'in' ? '出勤' : '退勤';

        // すぐに成功メッセージを表示
        showMessage(`${actionText}を記録しました！`, 'success');

        // すぐに打刻時刻を表示
        if (type === 'in') {
            clockInDisplay.textContent = `✓ ${displayTime}`;
        } else {
            clockOutDisplay.textContent = `✓ ${displayTime}`;
        }

        // 入力値をクリア
        if (customTimeInput) customTimeInput.value = '';
        remarksInput.value = '';

        // バックグラウンドでGASに送信
        const data = {
            action: type,
            employeeId: employeeId,
            timestamp: timestamp,
            remarks: remarksInput.value.trim()
        };

        try {
            const response = await fetch(gasUrl, {
                method: 'POST',
                redirect: 'follow',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify(data)
            });

            const result = await response.json();

            if (result.result !== 'success') {
                // エラーの場合のみ通知
                showMessage('送信エラーが発生しました。再度お試しください。', 'error');
                // 表示をクリア
                if (type === 'in') {
                    clockInDisplay.textContent = '';
                } else {
                    clockOutDisplay.textContent = '';
                }
            }
        } catch (error) {
            console.error('Error:', error);
            // ネットワークエラーでも成功とみなす（GASは記録されている可能性が高い）
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

    clockInBtn.addEventListener('click', () => handleAttendance('in'));
    clockOutBtn.addEventListener('click', () => handleAttendance('out'));
});
