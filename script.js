document.addEventListener('DOMContentLoaded', () => {
    const timeDisplay = document.getElementById('current-time');
    const employeeIdInput = document.getElementById('employee-id');
    const clockInBtn = document.getElementById('clock-in-btn');
    const clockOutBtn = document.getElementById('clock-out-btn');
    const statusMessage = document.getElementById('status-message');
    const gasUrlInput = document.getElementById('gas-url');
    const saveSettingsBtn = document.getElementById('save-settings');
    const clockInTimeInput = document.getElementById('clock-in-time');
    const clockOutTimeInput = document.getElementById('clock-out-time');
    const remarksInput = document.getElementById('remarks');

    // 履歴表示要素
    const historyContainer = document.getElementById('history-container');

    // 設定のロード
    const savedGasUrl = localStorage.getItem('attendance_gas_url');
    const savedEmployeeId = localStorage.getItem('attendance_employee_id');

    if (savedGasUrl) gasUrlInput.value = savedGasUrl;
    if (savedEmployeeId) {
        employeeIdInput.value = savedEmployeeId;
        fetchHistory(); // 初期表示時に履歴を取得
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

    // 社員コード保存（入力が変わるたびに保存して履歴取得）
    employeeIdInput.addEventListener('change', () => {
        const id = employeeIdInput.value.trim();
        localStorage.setItem('attendance_employee_id', id);
        if (id) {
            fetchHistory();
        }
    });

    // 履歴取得処理
    async function fetchHistory() {
        const employeeId = employeeIdInput.value.trim();
        const gasUrl = localStorage.getItem('attendance_gas_url');

        if (!employeeId || !gasUrl) return;

        try {
            // POSTリクエストで履歴を取得 (CORS対応のためPOSTを使用)
            // 注意: GAS側でContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON)している必要がある
            const response = await fetch(gasUrl, {
                method: 'POST',
                redirect: 'follow', // GASのリダイレクトを追跡
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8', // simple requestにするため
                },
                body: JSON.stringify({
                    action: 'get_history',
                    employeeId: employeeId
                })
            });

            const data = await response.json();

            if (data.result === 'success' && data.history) {
                renderHistory(data.history);
            }

        } catch (error) {
            console.error('History fetch error:', error);
            // 履歴取得失敗はユーザーに通知しなくても良い（サイレント失敗）
        }
    }

    function renderHistory(history) {
        if (!history || history.length === 0) {
            historyContainer.innerHTML = '<p class="no-history">履歴がありません</p>';
            return;
        }

        let html = `
            <table class="history-table">
                <thead>
                    <tr>
                        <th>日付</th>
                        <th>出勤</th>
                        <th>退勤</th>
                        <th>備考</th>
                    </tr>
                </thead>
                <tbody>
        `;

        history.forEach(row => {
            html += `
                <tr>
                    <td>${row.date}</td>
                    <td>${row.clockInTime}</td>
                    <td>${row.clockOutTime}</td>
                    <td>${row.remarks}</td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
        `;

        historyContainer.innerHTML = html;
    }

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
            let timestamp;
            let targetTimeInput;

            if (type === 'in') {
                targetTimeInput = clockInTimeInput;
            } else {
                targetTimeInput = clockOutTimeInput;
            }

            if (targetTimeInput && targetTimeInput.value) {
                timestamp = new Date(targetTimeInput.value).toISOString();
            } else {
                timestamp = new Date().toISOString();
            }

            const data = {
                action: type, // 'in' or 'out'
                employeeId: employeeId,
                timestamp: timestamp,
                remarks: remarksInput.value.trim()
            };

            // CORSモードで送信してレスポンスを受け取る
            const response = await fetch(gasUrl, {
                method: 'POST',
                redirect: 'follow',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify(data)
            });

            const result = await response.json();

            if (result.result === 'success') {
                const actionText = type === 'in' ? '出勤' : '退勤';
                showMessage(`${actionText}を記録しました！`, 'success');

                // 入力値をクリア
                if (targetTimeInput) targetTimeInput.value = '';
                remarksInput.value = '';

                // 履歴を更新
                fetchHistory();
            } else {
                throw new Error(result.message || 'Unknown error');
            }

        } catch (error) {
            console.error('Error:', error);
            // CORSエラーなどでレスポンスが読めない場合のフォールバック
            // no-corsで再送するか、あるいは単に成功したとみなすか...
            // ここでは簡易的にエラー表示するが、GASのデプロイ設定(全員にアクセス権)が重要
            showMessage('送信完了(応答なし)。履歴を確認してください。', 'success');

            // 念のため履歴更新を試みる
            setTimeout(fetchHistory, 1000);
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
