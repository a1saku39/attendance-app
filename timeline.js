document.addEventListener('DOMContentLoaded', () => {
    // Menu Logic
    const menuBtn = document.getElementById('menuBtn');
    const closeMenuBtn = document.getElementById('closeMenuBtn');
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('overlay');
    const statusMessage = document.getElementById('status-message');

    function toggleMenu() {
        if (sidebar && overlay) {
            sidebar.classList.toggle('active');
            overlay.classList.toggle('active');
        }
    }
    if (menuBtn) menuBtn.addEventListener('click', toggleMenu);
    if (closeMenuBtn) closeMenuBtn.addEventListener('click', toggleMenu);
    if (overlay) overlay.addEventListener('click', toggleMenu);

    // Timeline Logic
    const monthSelector = document.getElementById('month-selector');
    const timelineContainer = document.getElementById('timeline-container');
    const employeeId = localStorage.getItem('attendance_employee_id');
    const gasUrl = localStorage.getItem('attendance_gas_url');

    if (!employeeId || !gasUrl) {
        showMessage('設定（社員コード・GAS URL）が完了していません。ホーム画面から設定を行ってください。', 'error');
        timelineContainer.innerHTML = '<div style="text-align:center; padding: 20px;">設定が完了していません。<br><a href="index.html">ホームへ戻る</a></div>';
        return;
    }

    // Initialize Month Selector
    const today = new Date();
    const currentMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
    monthSelector.value = currentMonth;

    // Load Data
    loadTimelineData(currentMonth);

    monthSelector.addEventListener('change', (e) => {
        loadTimelineData(e.target.value);
    });

    async function loadTimelineData(yearMonth) {
        timelineContainer.innerHTML = '<div style="text-align: center; padding: 20px; color: #666;">読み込み中...</div>';

        try {
            const response = await fetch(gasUrl, {
                method: 'POST',
                redirect: 'follow',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    action: 'getPersonalMonthlyData',
                    employeeId: employeeId,
                    yearMonth: yearMonth
                })
            });

            const text = await response.text();
            let result;
            try {
                result = JSON.parse(text);
            } catch (e) {
                console.error('JSON Parse Error:', e);
                console.error('Raw Response:', text);
                throw new Error('サーバーからの応答が不正です (JSONではありません): ' + text.substring(0, 100));
            }

            if (result.result === 'success') {
                renderTimeline(result.data, yearMonth);
            } else {
                const errorMsg = result.message || '不明なエラー';
                showMessage('データの取得に失敗しました: ' + errorMsg, 'error');
                timelineContainer.innerHTML = `<div style="text-align: center; color: red;">データの取得に失敗しました<br>${errorMsg}</div>`;
            }
        } catch (error) {
            console.error('Error:', error);
            const errorMsg = error.message || error.toString();
            showMessage('エラーが発生しました: ' + errorMsg, 'error');
            timelineContainer.innerHTML = `<div style="text-align: center; color: red; padding: 20px; word-break: break-all;">
                <h3>エラーが発生しました</h3>
                <p>以下のエラー内容を確認してください:</p>
                <div style="background: #fee2e2; padding: 10px; border-radius: 4px; margin-top: 10px;">
                    ${errorMsg}
                </div>
            </div>`;
        }
    }

    function renderTimeline(data, yearMonth) {
        timelineContainer.innerHTML = '';

        const days = Object.keys(data).sort().reverse(); // Show newest first

        if (days.length === 0) {
            timelineContainer.innerHTML = '<div style="text-align:center; padding: 20px; color: #6b7280;">この月のデータはありません</div>';
            return;
        }

        days.forEach(dateStr => {
            const dayData = data[dateStr];
            // Skip empty days if needed, but usually we want to see even if just clocked in
            if (!dayData.clockInTime && !dayData.clockOutTime) return;

            const date = new Date(dateStr);
            const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
            const formattedDate = `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;

            const item = document.createElement('div');
            item.className = 'timeline-item';

            let contentHtml = '';

            // Check In
            if (dayData.clockInTime) {
                contentHtml += `
                    <div class="timeline-row">
                        <div style="display:flex; align-items:center; gap:10px;">
                            <span class="timeline-label label-in">出勤</span>
                            <span class="timeline-time">${dayData.clockInTime}</span>
                        </div>
                    </div>
                `;
            }

            // Check Out
            if (dayData.clockOutTime) {
                contentHtml += `
                    <div class="timeline-row">
                        <div style="display:flex; align-items:center; gap:10px;">
                            <span class="timeline-label label-out">退勤</span>
                            <span class="timeline-time">${dayData.clockOutTime}</span>
                        </div>
                    </div>
                `;
            }

            // Working Hours
            if (dayData.workingHours) {
                contentHtml += `
                    <div class="timeline-row" style="margin-top: 5px; font-size: 0.85rem; color: #666;">
                        <span style="background: #f3f4f6; padding: 2px 6px; border-radius: 4px;">時間: ${dayData.workingHours}</span>
                    </div>
                 `;
            }

            // Remarks
            if (dayData.remarks) {
                contentHtml += `<div class="timeline-remarks">${dayData.remarks}</div>`;
            }

            // GPS Badge
            const hasGps = dayData.locationLog && dayData.locationLog.length > 0;
            const gpsBadge = hasGps ? '<span style="background:#e0f2fe; color:#0284c7; font-size:0.75rem; padding:2px 6px; border-radius:4px; margin-left:8px; vertical-align:middle; font-weight:600;">GPS</span>' : '';

            item.innerHTML = `
                <div class="timeline-marker"></div>
                <div class="timeline-date">
                    ${formattedDate} 
                    <span class="timeline-day-week">(${dayOfWeek})</span>
                    ${gpsBadge}
                </div>
                <div class="timeline-content">
                    ${contentHtml}
                    <div id="map-${dateStr.replace(/\//g, '-')}" class="map-container" style="display:none; height: 300px; margin-top: 15px; border-radius: 8px;"></div>
                </div>
            `;

            timelineContainer.appendChild(item);

            // Render Map if location info exists
            if (dayData.locationLog && dayData.locationLog.length > 0) {
                const mapId = `map-${dateStr.replace(/\//g, '-')}`;
                const mapEl = document.getElementById(mapId);
                mapEl.style.display = 'block';

                setTimeout(() => {
                    initMap(mapId, dayData.locationLog);
                }, 100);
            }
        });
    }

    function initMap(elementId, locations) {
        if (!locations || locations.length === 0) return;

        // Create Leaflet Map
        // Default to first point
        const first = locations[0];
        const map = L.map(elementId).setView([first.lat, first.lng], 13);

        // Use OpenStreetMap tiles (Free, no key required)
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            maxZoom: 19,
            attribution: '© OpenStreetMap'
        }).addTo(map);

        // Add markers and path
        const latlngs = [];
        locations.forEach(loc => {
            if (loc.lat && loc.lng) {
                const marker = L.marker([loc.lat, loc.lng]).addTo(map);
                const time = loc.time || '';
                const action = loc.action === 'in' ? '出勤' : (loc.action === 'out' ? '退勤' : '記録');
                marker.bindPopup(`<b>${action}</b><br>${time}`);
                latlngs.push([loc.lat, loc.lng]);
            }
        });

        // If multiple points, fit bounds and draw line
        if (latlngs.length > 1) {
            const polyline = L.polyline(latlngs, { color: 'blue' }).addTo(map);
            map.fitBounds(polyline.getBounds(), { padding: [50, 50] });
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
});
