// ローカルストレージを使用してタスクを管理
let tasks = JSON.parse(localStorage.getItem('tasks')) || [];

// タスクを表示
function renderTasks() {
    const taskTableBody = document.getElementById('taskTableBody');
    taskTableBody.innerHTML = '';
    tasks.forEach((task, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><input type="checkbox" class="delete-checkbox" data-index="${index}"></td>
            <td>${index + 1}</td>
            <td>${task.name}</td>
            <td>${task.memo}</td>
            <td><a href="${task.link}" target="_blank">${task.link}</a></td>
            <td>${task.startDate}</td>
            <td>${task.dueDate}</td>
        `;
        taskTableBody.appendChild(row);
    });
}

// タスク追加
document.getElementById('addTaskButton').addEventListener('click', () => {
    const taskName = document.getElementById('taskName').value;
    const memo = document.getElementById('memo').value;
    const relatedLink = document.getElementById('relatedLink').value;
    const startDate = document.getElementById('startDate').value;
    const dueDate = document.getElementById('dueDate').value;

    if (taskName && dueDate) {
        tasks.push({ name: taskName, memo, link: relatedLink, startDate, dueDate });
        localStorage.setItem('tasks', JSON.stringify(tasks));
        renderTasks();
        alert('タスクが登録されました');
    } else {
        alert('Task名と期日は必須です');
    }
});

// タスク削除
document.getElementById('deleteTaskButton').addEventListener('click', () => {
    // 削除対象のタスクを収集
    const checkboxes = document.querySelectorAll('.delete-checkbox:checked');
    const indexesToDelete = Array.from(checkboxes).map(checkbox => Number(checkbox.dataset.index));

    // インデックス順に降順で削除（降順にしないと配列操作でズレる）
    indexesToDelete.sort((a, b) => b - a).forEach(index => tasks.splice(index, 1));

    // ローカルストレージを更新
    localStorage.setItem('tasks', JSON.stringify(tasks));

    // 画面を更新
    renderTasks();
    alert('選択したタスクを削除しました');
});

// ポップアップ通知
function showPopup(task) {
    const popup = document.createElement('div');
    popup.className = 'popup';
    popup.textContent = `タスク: ${task.name} | 期日: ${task.dueDate}`;
    document.body.appendChild(popup);

    setTimeout(() => {
        popup.remove();
    }, 5000);
}

// 期日チェック
function checkDeadlines() {
    const today = new Date();
    tasks.forEach(task => {
        const dueDate = new Date(task.dueDate);
        const diff = (dueDate - today) / (1000 * 60 * 60 * 24);
        if (diff <= 2 && diff > 0) {
            showPopup(task);
        }
    });
}

// 初期化
renderTasks();
checkDeadlines();
