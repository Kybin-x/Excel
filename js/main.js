/* ============================================
   Excel学习平台 - 核心脚本
   ============================================ */

// ========== 导航菜单 ==========
document.getElementById('menuToggle')?.addEventListener('click', function () {
    document.getElementById('navLinks').classList.toggle('show');
});

// 点击页面其他区域关闭菜单
document.addEventListener('click', function (e) {
    const nav = document.getElementById('navLinks');
    const toggle = document.getElementById('menuToggle');
    if (nav && toggle && !nav.contains(e.target) && !toggle.contains(e.target)) {
        nav.classList.remove('show');
    }
});

// ========== 手风琴（课次展开/收起） ==========
function toggleLesson(lessonId) {
    const card = document.getElementById('lesson-card-' + lessonId);
    const body = document.getElementById('lesson-body-' + lessonId);
    if (!card || !body) return;

    const isActive = card.classList.contains('active');

    // 关闭其他
    document.querySelectorAll('.lesson-card').forEach(c => {
        c.classList.remove('active');
        c.querySelector('.lesson-body').style.maxHeight = '0';
    });

    // 切换当前
    if (!isActive) {
        card.classList.add('active');
        body.style.maxHeight = body.scrollHeight + 'px';
        // 滚动到视图
        setTimeout(() => {
            card.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
    }
}

// ========== 测验系统 ==========
function renderQuiz(containerId, questions) {
    const container = document.getElementById(containerId);
    if (!container) return;

    let html = '<div class="quiz-questions">';
    questions.forEach((q, i) => {
        html += `<div class="quiz-question" id="${containerId}-q${i}">`;
        html += `<div class="quiz-q-text">${i + 1}. ${q.q}</div>`;
        html += '<div class="quiz-options">';
        const labels = ['A', 'B', 'C', 'D'];
        q.opts.forEach((opt, j) => {
            html += `<label id="${containerId}-q${i}-opt${j}">`;
            html += `<input type="radio" name="${containerId}-q${i}" value="${j}"> `;
            html += `${labels[j]}. ${opt}</label>`;
        });
        html += '</div>';
        html += `<div class="quiz-explain" id="${containerId}-q${i}-explain">💡 ${q.explain || ''}</div>`;
        html += '</div>';
    });
    html += '</div>';
    html += `<div style="text-align:center;margin-top:16px;">
        <button class="btn-quiz" onclick="submitQuiz('${containerId}')">📝 提交答案</button>
    </div>`;
    html += `<div class="quiz-result" id="${containerId}-result" style="display:none;"></div>`;

    container.innerHTML = html;
}

function submitQuiz(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const questions = container._quizData;
    if (!questions) return;

    let score = 0;
    let unanswered = 0;

    questions.forEach((q, i) => {
        const qDiv = document.getElementById(`${containerId}-q${i}`);
        const selected = container.querySelector(`input[name="${containerId}-q${i}"]:checked`);
        const explainDiv = document.getElementById(`${containerId}-q${i}-explain`);

        if (!selected) {
            unanswered++;
            return;
        }

        const selectedVal = parseInt(selected.value);
        const labels = ['A', 'B', 'C', 'D'];

        if (selectedVal === q.ans) {
            score++;
            qDiv.classList.add('correct');
            qDiv.classList.remove('wrong');
        } else {
            qDiv.classList.add('wrong');
            qDiv.classList.remove('correct');
            // 高亮正确选项
            const correctLabel = document.getElementById(`${containerId}-q${i}-opt${q.ans}`);
            if (correctLabel) correctLabel.classList.add('opt-correct');
            // 标记错误选项
            const wrongLabel = document.getElementById(`${containerId}-q${i}-opt${selectedVal}`);
            if (wrongLabel) wrongLabel.classList.add('opt-wrong');
        }

        if (explainDiv && q.explain) {
            explainDiv.style.display = 'block';
        }
    });

    if (unanswered > 0) {
        alert(`还有 ${unanswered} 题未作答，请完成所有题目！`);
        return;
    }

    // 禁用所有radio
    container.querySelectorAll('input[type="radio"]').forEach(r => r.disabled = true);

    // 显示结果
    const resultDiv = document.getElementById(`${containerId}-result`);
    let msg = '', emoji = '';
    if (score === 10) { emoji = '🎉'; msg = '太棒了！满分！你已经完全掌握了这部分知识！'; }
    else if (score >= 8) { emoji = '👏'; msg = '优秀！掌握得很好，再巩固一下就完美了！'; }
    else if (score >= 6) { emoji = '💪'; msg = '不错！大部分知识已经掌握，继续加油！'; }
    else if (score >= 4) { emoji = '📖'; msg = '还需要多复习，建议重新看看知识讲解！'; }
    else { emoji = '😊'; msg = '别灰心！回头认真看看知识点，再来挑战！'; }

    resultDiv.innerHTML = `
        <div class="quiz-score" style="color:${score >= 6 ? 'var(--green)' : 'var(--orange)'}">${score}/10</div>
        <div class="quiz-msg">${emoji} ${msg}</div>
        <div class="quiz-actions">
            <button class="btn-retry" onclick="retryQuiz('${containerId}')">🔄 重新测试</button>
        </div>
    `;
    resultDiv.style.display = 'block';

    // 保存进度
    saveQuizScore(containerId, score);

    resultDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function retryQuiz(containerId) {
    const container = document.getElementById(containerId);
    if (!container || !container._quizData) return;

    // 重置所有状态
    container.querySelectorAll('.quiz-question').forEach(q => {
        q.classList.remove('correct', 'wrong');
    });
    container.querySelectorAll('label').forEach(l => {
        l.classList.remove('opt-correct', 'opt-wrong');
    });
    container.querySelectorAll('input[type="radio"]').forEach(r => {
        r.checked = false;
        r.disabled = false;
    });
    container.querySelectorAll('.quiz-explain').forEach(e => {
        e.style.display = 'none';
    });
    const resultDiv = document.getElementById(`${containerId}-result`);
    if (resultDiv) resultDiv.style.display = 'none';
}

function initQuiz(containerId, questions) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container._quizData = questions;
    renderQuiz(containerId, questions);
}

// ========== 进度管理 ==========
function saveQuizScore(quizId, score) {
    try {
        const progress = JSON.parse(localStorage.getItem('excelProgress') || '{}');
        progress[quizId] = { score: score, date: new Date().toISOString() };
        localStorage.setItem('excelProgress', JSON.stringify(progress));
        updateProgressBar();
    } catch (e) { /* ignore */ }
}

function getQuizScore(quizId) {
    try {
        const progress = JSON.parse(localStorage.getItem('excelProgress') || '{}');
        return progress[quizId]?.score ?? null;
    } catch (e) { return null; }
}

function updateProgressBar() {
    const bar = document.getElementById('progressFill');
    const text = document.getElementById('progressText');
    if (!bar || !text) return;

    const totalLessons = parseInt(bar.dataset.total || '5');
    let completed = 0;
    for (let i = 1; i <= totalLessons; i++) {
        const moduleId = bar.dataset.module || '1';
        const score = getQuizScore(`quiz-${moduleId}-${i}`);
        if (score !== null && score >= 6) completed++;
    }
    const percent = Math.round((completed / totalLessons) * 100);
    bar.style.width = percent + '%';
    text.innerHTML = `已完成 <span>${completed}/${totalLessons}</span> 课 (${percent}%)`;
}

// ========== CSV下载 ==========
function downloadCSV(filename, headers, rows) {
    const BOM = '\uFEFF';
    let csv = BOM;
    csv += headers.join(',') + '\n';
    rows.forEach(row => {
        csv += row.map(cell => {
            // 如果包含逗号或引号，用双引号包裹
            const str = String(cell);
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        }).join(',') + '\n';
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

// ========== 初始化 ==========
document.addEventListener('DOMContentLoaded', function () {
    updateProgressBar();
});
