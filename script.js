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

    // バージョンチェック
    checkVersion();

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
            checkAdmin(id);
        } else {
            hideAdminMenus();
        }
    });

    // 管理者チェック
    async function checkAdmin(empId) {
        const gasUrl = localStorage.getItem('attendance_gas_url');
        if (!gasUrl || !empId) return;

        try {
            const response = await fetch(gasUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    action: 'checkAdminStatus',
                    employeeId: empId
                })
            });
            const result = await response.json();
            if (result.result === 'success' && result.isAdmin) {
                document.getElementById('menu-survey-admin').style.display = 'block';
                document.getElementById('menu-proposal-admin').style.display = 'block';
            } else {
                hideAdminMenus();
            }
        } catch (e) {
            hideAdminMenus();
        }
    }

    function hideAdminMenus() {
        const menu1 = document.getElementById('menu-survey-admin');
        const menu2 = document.getElementById('menu-proposal-admin');
        if(menu1) menu1.style.display = 'none';
        if(menu2) menu2.style.display = 'none';
    }

    // 初回ロード時
    if (savedEmployeeId) {
        checkAdmin(savedEmployeeId);
    }

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

        // 備考の内容を保持しておく
        const remarksVal = remarksInput.value.trim();

        // 入力値をクリア
        if (customDateInput) customDateInput.value = '';
        if (customTimeInput) customTimeInput.value = '';
        remarksInput.value = '';

        // バックグラウンドでGASに送信
        const data = {
            action: type,
            employeeId: employeeId,
            timestamp: timestamp,
            remarks: remarksVal,
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

    // 未読メッセージチェック
    checkUnreadMessages();
    // 3分ごとに未読チェック
    setInterval(checkUnreadMessages, 3 * 60 * 1000);

    // 掲示板バナーの読み込み
    loadBoardBanner();
});

// バージョンチェック機能
async function checkVersion() {
    const gasUrl = localStorage.getItem('attendance_gas_url');
    if (!gasUrl) return; // URLが設定されていない場合はスキップ

    const currentVersion = localStorage.getItem('attendance_app_version') || 'v3.6';

    // まずlocalStorageのバージョンを即表示
    updateMenuVersion(currentVersion);

    try {
        const response = await fetch(gasUrl, {
            method: 'POST',
            redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'getVersionInfo' })
        });

        const text = await response.text();
        const data = JSON.parse(text);

        if (data.result === 'success' && data.versionInfo) {
            const serverVersion = data.versionInfo.version;

            // メニューのバージョン表示を更新
            updateMenuVersion(serverVersion);

            // 新しいバージョンがある場合
            if (serverVersion !== currentVersion) {
                showVersionUpdateNotification(data.versionInfo);
                // バージョンを保存
                localStorage.setItem('attendance_app_version', serverVersion);
            }
        }
    } catch (error) {
        console.log('バージョンチェックエラー:', error);
        // エラーは無視（ユーザーに影響を与えない）
    }
}

// メニューのバージョン表示を更新
function updateMenuVersion(version) {
    const versionEl = document.getElementById('menuVersion');
    if (versionEl && version) {
        versionEl.textContent = version;
    }
}

function showVersionUpdateNotification(versionInfo) {
    const statusMessage = document.getElementById('status-message');
    if (!statusMessage) return;

    let notesHtml = '';
    if (versionInfo.updateNotes && versionInfo.updateNotes.length > 0) {
        notesHtml = '<ul style="text-align: left; margin: 10px 0; padding-left: 20px;">';
        versionInfo.updateNotes.forEach(note => {
            notesHtml += `<li>${note}</li>`;
        });
        notesHtml += '</ul>';
    }

    const message = `
        <div style="padding: 15px; background: #e3f2fd; border: 2px solid #2196f3; border-radius: 8px; margin-bottom: 15px;">
            <div style="font-weight: bold; color: #1976d2; margin-bottom: 8px;">
                🎉 アプリが更新されました！ ${versionInfo.version}
            </div>
            <div style="font-size: 13px; color: #555;">
                リリース日: ${versionInfo.releaseDate}
            </div>
            ${notesHtml}
            <button onclick="this.parentElement.remove()" 
                    style="margin-top: 10px; padding: 6px 15px; background: #2196f3; color: white; border: none; border-radius: 4px; cursor: pointer;">
                閉じる
            </button>
        </div>
    `;

    // アプリのコンテナ（.container）内に通知を挿入して、スマホで右側にはみ出さないようにする
    const appContainer = document.querySelector('.container');
    const mainArea = document.querySelector('main');

    const notification = document.createElement('div');
    notification.innerHTML = message;

    if (appContainer && mainArea) {
        appContainer.insertBefore(notification, mainArea);
    } else if (appContainer) {
        appContainer.insertBefore(notification, appContainer.firstChild);
    } else {
        const container = statusMessage.parentElement;
        container.insertBefore(notification, statusMessage);
    }
}

// 未読メッセージチェック機能
async function checkUnreadMessages() {
    const gasUrl = localStorage.getItem('attendance_gas_url');
    const employeeCode = localStorage.getItem('attendance_employee_id');

    if (!gasUrl || !employeeCode) return;

    try {
        const response = await fetch(gasUrl, {
            method: 'POST',
            redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                action: 'getMessages',
                employeeCode: employeeCode
            })
        });

        const text = await response.text();
        const data = JSON.parse(text);

        if (data.result === 'success') {
            const received = data.received || [];
            const unreadCount = received.filter(function (m) { return m.status === '未読'; }).length;

            // メニュー内のバッジを更新
            const badge = document.getElementById('msgUnreadBadge');
            if (badge) {
                if (unreadCount > 0) {
                    badge.textContent = unreadCount;
                    badge.style.display = 'inline-block';
                } else {
                    badge.style.display = 'none';
                }
            }

            // メニューボタンのドットを更新
            const dot = document.getElementById('menuDot');
            if (dot) {
                dot.style.display = unreadCount > 0 ? 'block' : 'none';
            }
        }
    } catch (error) {
        console.error('未読チェックエラー:', error);
    }
}

// ========================================
// 掲示板バナー
// ========================================
async function loadBoardBanner() {
    const gasUrl = localStorage.getItem('attendance_gas_url');
    const bannerEl = document.getElementById('boardBanner');
    if (!gasUrl || !bannerEl) return;

    try {
        const response = await fetch(gasUrl, {
            method: 'POST',
            redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'getBoardPosts' })
        });
        const text = await response.text();
        let data;
        try { data = JSON.parse(text); } catch (e) { return; }

        if (data.result !== 'success' || !data.posts || data.posts.length === 0) return;

        // 閉じた投稿IDを取得 (タイトル+日付のハッシュで管理)
        const dismissed = JSON.parse(localStorage.getItem('board_dismissed') || '[]');

        // 直近30日・最大5件に絞り込み
        const cutoff = Date.now() - 30 * 24 * 60 * 60 * 1000;
        const typeIcons = { important: '🔔', update: '🆕', info: 'ℹ️' };
        const typeClass = { important: 'type-important', update: 'type-update', info: 'type-info' };

        let html = '';
        let shown = 0;

        for (const post of data.posts) {
            if (shown >= 5) break;
            const key = (post.title + post.date).replace(/\s/g, '');
            if (dismissed.includes(key)) continue;

            const cls = typeClass[post.type] || 'type-default';
            const icon = typeIcons[post.type] || '📌';
            // 本文プレビュー（改行をスペースに変換）
            const preview = (post.content || '').replace(/\\n/g, ' ').replace(/\n/g, ' ');

            html += `
                <div class="board-banner-item ${cls}" id="banner-${shown}">
                    <span class="board-banner-icon">${icon}</span>
                    <div class="board-banner-body">
                        <div class="board-banner-header">
                            <span class="board-banner-title">${escBanner(post.title)}</span>
                            <span class="board-banner-date">${post.date}</span>
                        </div>
                        <div class="board-banner-preview">${escBanner(preview)}</div>
                        <a href="Board.html" class="board-banner-link">📌 掲示板を開く →</a>
                    </div>
                    <button class="board-banner-close" onclick="dismissBanner('${key}', 'banner-${shown}')" title="閉じる">✕</button>
                </div>`;
            shown++;
        }

        if (shown > 0) {
            bannerEl.innerHTML = html;
            bannerEl.style.display = 'block';
        }

    } catch (e) {
        console.error('掲示板バナー取得エラー:', e);
    }
}

function dismissBanner(key, elemId) {
    // 要素を削除
    const el = document.getElementById(elemId);
    if (el) {
        el.style.opacity = '0';
        el.style.transform = 'translateY(-6px)';
        el.style.transition = 'opacity 0.25s, transform 0.25s';
        setTimeout(() => {
            el.remove();
            // すべて閉じたらバナーエリアを非表示
            const bannerEl = document.getElementById('boardBanner');
            if (bannerEl && bannerEl.children.length === 0) {
                bannerEl.style.display = 'none';
            }
        }, 250);
    }
    // 閉じた項目を記録（最大50件まで保持）
    const dismissed = JSON.parse(localStorage.getItem('board_dismissed') || '[]');
    if (!dismissed.includes(key)) {
        dismissed.push(key);
        if (dismissed.length > 50) dismissed.shift();
        localStorage.setItem('board_dismissed', JSON.stringify(dismissed));
    }
}

function escBanner(str) {
    const d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
}

// サイドメニューのアコーディオン開閉
function toggleMenuGroup(element) {
    const group = element.closest('.menu-group');
    if (group) {
        group.classList.toggle('active');
    }
}
