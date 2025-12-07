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

                let isClockedIn = false;
                let isClockedOut = false;

                if (todayData) {
                    if (todayData.clockInTime) {
                        clockInDisplay.textContent = `✓ ${todayData.clockInTime}`;
                        isClockedIn = true;
                    }
                    if (todayData.clockOutTime) {
                        clockOutDisplay.textContent = `✓ ${todayData.clockOutTime}`;
                        isClockedOut = true;
                    }
                }

                // 打刻忘れチェックと警告音
                checkAttendanceAlert(isClockedIn, isClockedOut);
            }
        } catch (error) {
            console.error('本日の打刻時刻取得エラー:', error);
        }
    }

    // 警告音を鳴らす関数 (Web Audio API)
    function playAlertSound() {
        try {
            const AudioContext = window.AudioContext || window.webkitAudioContext;
            if (!AudioContext) return;

            const audioCtx = new AudioContext();
            const oscillator = audioCtx.createOscillator();
            const gainNode = audioCtx.createGain();

            oscillator.type = 'square'; // 矩形波（警告音っぽい音）

            // ピッ・ピッ・ピッ というパターン
            const now = audioCtx.currentTime;

            oscillator.frequency.setValueAtTime(880, now); // 880Hz (ラ)
            oscillator.frequency.setValueAtTime(880, now + 0.1);
            oscillator.frequency.setValueAtTime(0, now + 0.1); // 無音

            oscillator.frequency.setValueAtTime(880, now + 0.2);
            oscillator.frequency.setValueAtTime(880, now + 0.3);
            oscillator.frequency.setValueAtTime(0, now + 0.3);

            oscillator.frequency.setValueAtTime(880, now + 0.4);
            oscillator.frequency.setValueAtTime(880, now + 0.5);

            gainNode.gain.setValueAtTime(0.1, now); // 音量 10%
            gainNode.gain.exponentialRampToValueAtTime(0.001, now + 0.5);

            oscillator.connect(gainNode);
            gainNode.connect(audioCtx.destination);

            oscillator.start();
            oscillator.stop(now + 0.6);
        } catch (e) {
            console.error('警告音再生エラー:', e);
        }
    }

    // 打刻忘れチェック関数
    function checkAttendanceAlert(isClockedIn, isClockedOut) {
        const now = new Date();
        const hour = now.getHours();

        // 土日はスキップ
        const day = now.getDay();
        if (day === 0 || day === 6) return;

        let shouldAlert = false;
        let alertMessage = '';

        // 条件1: 12時を過ぎて出勤していない場合
        if (hour >= 12 && !isClockedIn) {
            shouldAlert = true;
            alertMessage = '出勤打刻がされていません！';
        }
        // 条件2: 18時を過ぎて出勤済みだが退勤していない場合
        else if (hour >= 18 && isClockedIn && !isClockedOut) {
            shouldAlert = true;
            alertMessage = '退勤打刻がされていません！';
        }

        if (shouldAlert) {
            // 画面にメッセージ表示
            showMessage(alertMessage, 'error');

            // 音を鳴らす (ユーザー操作が必要な場合があるため、try-catchで囲む)
            // ※ブラウザのポリシーにより、ユーザーが一度でもページを操作していないと音は鳴りません
            playAlertSound();
        }
    }

    // 定期的に打刻状況を再チェック (5分ごと)
    setInterval(() => {
        const employeeId = employeeIdInput.value.trim();
        if (employeeId) {
            loadTodayAttendance();
        }
    }, 5 * 60 * 1000);

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

// ハンバーガーメニュー制御
document.addEventListener('DOMContentLoaded', function () {
    const menuBtn = document.getElementById('menuBtn');
    const closeMenuBtn = document.getElementById('closeMenuBtn');
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('overlay');

    function toggleMenu() {
        if (sidebar && overlay) {
            sidebar.classList.toggle('active');
            overlay.classList.toggle('active');
        }
    }

    if (menuBtn) {
        menuBtn.addEventListener('click', toggleMenu);
    }

    if (closeMenuBtn) {
        closeMenuBtn.addEventListener('click', toggleMenu);
    }

    if (overlay) {
        overlay.addEventListener('click', toggleMenu);
    }
});
