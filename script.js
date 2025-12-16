document.addEventListener('DOMContentLoaded', () => {
    const timeDisplay = document.getElementById('current-time');
    const employeeIdInput = document.getElementById('employee-id');
    const clockInBtn = document.getElementById('clock-in-btn');
    const clockOutBtn = document.getElementById('clock-out-btn');
    const statusMessage = document.getElementById('status-message');
    const gasUrlInput = document.getElementById('gas-url');
    const saveSettingsBtn = document.getElementById('save-settings');
    const customDateInput = document.getElementById('custom-date');
    const customTimeInput = document.getElementById('custom-time');
    const remarksInput = document.getElementById('remarks');
    const clockInDisplay = document.getElementById('clock-in-display');
    const clockOutDisplay = document.getElementById('clock-out-display');
    const holidayBtn = document.getElementById('holiday-btn');
    const holidayDisplay = document.getElementById('holiday-display');
    const holidayModal = document.getElementById('holiday-modal');
    const paidLeaveBtn = document.getElementById('paid-leave-btn');
    const compensatoryLeaveBtn = document.getElementById('compensatory-leave-btn');
    const cancelHolidayBtn = document.getElementById('cancel-holiday-btn');

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

    // 打刻処理（楽観的UI更新 + GPS取得）
    async function handleAttendance(type, option = null) {
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

        // 位置情報取得機能は無効化されています
        // showMessage('位置情報を取得中...', 'success');

        let locationData = null;
        // try {
        //     const position = await getCurrentPosition();
        //     locationData = {
        //         lat: position.coords.latitude,
        //         lng: position.coords.longitude
        //     };
        // } catch (e) {
        //     console.warn('位置情報の取得に失敗しました:', e);
        //     showMessage('位置情報の取得に失敗しました。このまま記録します。', 'error');
        //     // 位置情報なしでも続行
        // }

        // 楽観的UI更新
        let timestamp;

        // 日付と時刻の指定がある場合の処理
        const dateVal = customDateInput ? customDateInput.value : '';
        const timeVal = customTimeInput ? customTimeInput.value : '';

        if (dateVal || timeVal) {
            const now = new Date();
            let year, month, day, hour, minute, second;

            // 日付の決定
            if (dateVal) {
                const dateParts = dateVal.split('-');
                year = parseInt(dateParts[0]);
                month = parseInt(dateParts[1]) - 1;
                day = parseInt(dateParts[2]);
            } else {
                year = now.getFullYear();
                month = now.getMonth();
                day = now.getDate();
            }

            // 時刻の決定
            if (timeVal) {
                const timeParts = timeVal.split(':');
                hour = parseInt(timeParts[0]);
                minute = parseInt(timeParts[1]);
                second = 0;
            } else {
                hour = now.getHours();
                minute = now.getMinutes();
                second = now.getSeconds();
            }

            timestamp = new Date(year, month, day, hour, minute, second).toISOString();
        } else {
            timestamp = new Date().toISOString();
        }

        const displayTime = new Date(timestamp).toLocaleTimeString('ja-JP', {
            hour: '2-digit',
            minute: '2-digit'
        });

        let actionText = '';
        if (type === 'in') actionText = '出勤';
        else if (type === 'out') actionText = '退勤';
        else if (type === 'holiday') actionText = option === 'paid_leave' ? '有給休暇' : '代休';

        // 位置情報取得完了後のメッセージ
        showMessage(`${actionText}を記録しました！`, 'success');

        // すぐに打刻時刻を表示
        if (type === 'in') {
            clockInDisplay.textContent = `✓ ${displayTime}`;
        } else if (type === 'out') {
            clockOutDisplay.textContent = `✓ ${displayTime}`;
        } else if (type === 'holiday') {
            holidayDisplay.textContent = `✓ ${actionText}`;
            // 休日設定時は他をクリアすべきか？とりあえずそのまま
        }

        // 入力値をクリア
        if (customDateInput) customDateInput.value = '';
        if (customTimeInput) customTimeInput.value = '';
        remarksInput.value = '';

        // バックグラウンドでGASに送信
        const data = {
            action: type,
            employeeId: employeeId,
            timestamp: timestamp,
            remarks: remarksInput.value.trim(),
            remarks: remarksInput.value.trim(),
            location: locationData, // 位置情報を追加
            option: option // 休日種別 (optional)
        };

        try {
            const response = await fetch(gasUrl, {
                method: 'POST',
                redirect: 'follow', // エラー対策
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify(data)
            });

            const result = await response.json(); // ここはテキスト→JSONの厳密なチェックは省略（成功率優先）

            if (result.result !== 'success') {
                showMessage('送信エラーが発生しました。再度お試しください。', 'error');
                if (type === 'in') {
                    clockInDisplay.textContent = '';
                } else if (type === 'out') {
                    clockOutDisplay.textContent = '';
                } else if (type === 'holiday') {
                    holidayDisplay.textContent = '';
                }
            }
        } catch (error) {
            console.error('Error:', error);
        }
    }

    // 位置情報取得のヘルパー関数（無効化されています）
    // function getCurrentPosition() {
    //     return new Promise((resolve, reject) => {
    //         if (!navigator.geolocation) {
    //             reject(new Error('Geolocation is not supported by this browser.'));
    //             return;
    //         }
    //         navigator.geolocation.getCurrentPosition(resolve, reject, {
    //             enableHighAccuracy: true,
    //             timeout: 10000,
    //             maximumAge: 0
    //         });
    //     });
    // }

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

    // 休日ボタン
    if (holidayBtn && holidayModal) {
        holidayBtn.addEventListener('click', () => {
            holidayModal.style.display = 'flex';
            setTimeout(() => holidayModal.classList.add('active'), 10);
        });

        cancelHolidayBtn.addEventListener('click', () => {
            holidayModal.classList.remove('active');
            setTimeout(() => holidayModal.style.display = 'none', 300);
        });

        paidLeaveBtn.addEventListener('click', () => {
            const currentRemarks = remarksInput.value.trim();
            const textToAdd = "【有給休暇】";
            if (!currentRemarks.includes(textToAdd)) {
                remarksInput.value = textToAdd + (currentRemarks ? " " + currentRemarks : "");
            }
            handleAttendance('holiday', 'paid_leave');
            holidayModal.classList.remove('active');
            setTimeout(() => holidayModal.style.display = 'none', 300);
        });

        compensatoryLeaveBtn.addEventListener('click', () => {
            const currentRemarks = remarksInput.value.trim();
            const textToAdd = "【代休】";
            if (!currentRemarks.includes(textToAdd)) {
                remarksInput.value = textToAdd + (currentRemarks ? " " + currentRemarks : "");
            }
            handleAttendance('holiday', 'compensatory');
            holidayModal.classList.remove('active');
            setTimeout(() => holidayModal.style.display = 'none', 300);
        });
    }

    // --- 定期的な位置情報記録 (無効化されています) ---
    // function sendLocationLog() {
    //     const employeeId = employeeIdInput.value.trim();
    //     const gasUrl = localStorage.getItem('attendance_gas_url');
    //
    //     if (!employeeId || !gasUrl) return;
    //
    //     getCurrentPosition().then(async (position) => {
    //         const locationData = {
    //             lat: position.coords.latitude,
    //             lng: position.coords.longitude
    //         };
    //
    //         const timestamp = new Date().toISOString();
    //         const data = {
    //             action: 'location',
    //             employeeId: employeeId,
    //             timestamp: timestamp,
    //             location: locationData
    //         };
    //
    //         console.log('Sending periodic location log...', data);
    //
    //         try {
    //             await fetch(gasUrl, {
    //                 method: 'POST',
    //                 redirect: 'follow',
    //                 headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    //                 body: JSON.stringify(data)
    //             });
    //             console.log('Location log sent successfully.');
    //         } catch (e) {
    //             console.error('Failed to send location log:', e);
    //         }
    //     }).catch(err => {
    //         console.warn('Periodic location check failed:', err);
    //     });
    // }
    //
    // // 1時間 = 60分 * 60秒 * 1000ミリ秒
    // const ONE_HOUR = 60 * 60 * 1000;
    //
    // // ページを開いてから即座に一度送るか、1時間後か？
    // // 「1時間ごと」なので、まずはインターバルをセット
    // setInterval(sendLocationLog, ONE_HOUR);
    //
    // // ページ読み込み時にも一度送信（オプション）
    // // ユーザー体験を損なわないよう少し遅らせて実行
    // setTimeout(sendLocationLog, 5000);

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
