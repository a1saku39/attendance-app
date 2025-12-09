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
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    action: 'getPersonalMonthlyData',
                    employeeId: employeeId,
                    yearMonth: yearMonth
                })
            });

            const result = await response.json();

            if (result.result === 'success') {
                renderTimeline(result.data, yearMonth);
            } else {
                showMessage('データの取得に失敗しました: ' + result.message, 'error');
                timelineContainer.innerHTML = '<div style="text-align: center; color: red;">データの取得に失敗しました</div>';
            }
        } catch (error) {
            console.error('Error:', error);
            showMessage('通信エラーが発生しました', 'error');
            timelineContainer.innerHTML = '<div style="text-align: center; color: red;">通信エラーが発生しました</div>';
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
            const formattedDate = `${date.getMonth() + 1}/${date.getDate()}`;

            const item = document.createElement('div');
            item.className = 'timeline-item';

            let contentHtml = '';

            // Check In
            if (dayData.clockInTime) {
                contentHtml += `
                    <div class="timeline-row">
                        <span class="timeline-label label-in">出勤</span>
                        <span class="timeline-time">${dayData.clockInTime}</span>
                    </div>
                `;
            }

            // Check Out
            if (dayData.clockOutTime) {
                contentHtml += `
                    <div class="timeline-row">
                        <span class="timeline-label label-out">退勤</span>
                        <span class="timeline-time">${dayData.clockOutTime}</span>
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

            item.innerHTML = `
                <div class="timeline-marker"></div>
                <div class="timeline-date">
                    ${formattedDate} 
                    <span class="timeline-day-week">(${dayOfWeek})</span>
                </div>
                <div class="timeline-content">
                    ${contentHtml}
                </div>
            `;

            timelineContainer.appendChild(item);
        });
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
