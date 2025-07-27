// DOM要素の取得
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const loading = document.getElementById('loading');
const error = document.getElementById('error');
const templatesContainer = document.getElementById('templatesContainer');

// LINE送信状態を管理するオブジェクト
let lineSentStatus = {};

// イベントリスナーの設定
document.addEventListener('DOMContentLoaded', function() {
    // アップロードエリアのイベント
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    
    // ファイル選択のイベント
    fileInput.addEventListener('change', handleFileSelect);
});

// ドラッグオーバー処理
function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
}

// ドラッグリーブ処理
function handleDragLeave(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
}

// ドロップ処理
function handleDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

// ファイル選択処理
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

// エラー表示
function showError(message) {
    error.textContent = message;
    error.style.display = 'block';
    loading.style.display = 'none';
}

// エラー非表示
function hideError() {
    error.style.display = 'none';
}

// ファイル処理のメイン関数
async function processFile(file) {
    // ファイル形式チェック
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Excelファイル（.xlsx または .xls）を選択してください');
        return;
    }

    hideError();
    loading.style.display = 'block';
    templatesContainer.style.display = 'none';

    // LINE送信状態をリセット
    lineSentStatus = {};

    try {
        // ファイル読み込み
        const data = await readFile(file);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // 特訓リストシートの存在チェック
        if (!workbook.Sheets['特訓リスト']) {
            throw new Error('「特訓リスト」シートが見つかりません。シート名を確認してください。');
        }

        const worksheet = workbook.Sheets['特訓リスト'];
        
        // データを JSON に変換（3行目から開始）
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            range: 2, // 3行目から開始（0ベース）
            defval: ''
        });

        if (jsonData.length === 0) {
            throw new Error('データが見つかりません。ファイルの内容を確認してください。');
        }

        // ヘッダー行を除いた実際のデータ
        const actualData = jsonData.slice(1);

        if (actualData.length === 0) {
            throw new Error('データ行が見つかりません。');
        }

        // データを適切な形式に変換
        const properData = actualData.map(row => ({
            'チェック': row['__EMPTY'] || '',
            'No': row['__EMPTY_1'] || '',
            'reg開始': row['__EMPTY_2'] || '',
            '指導開始': row['__EMPTY_3'] || '',
            '指導終了': row['__EMPTY_4'] || '',
            '生徒氏名': row['__EMPTY_5'] || '',
            '担当': row['__EMPTY_6'] || '',
            '形式': row['__EMPTY_7'] || '',
            '科目': row['__EMPTY_8'] || '',
            '状態': row['__EMPTY_9'] || '',
            '学年': row['__EMPTY_10'] || '',
            '校舎': row['__EMPTY_11'] || '',
            '基本': row['__EMPTY_12'] || '',
            '件名': row['__EMPTY_13'] || ''
        }));

        // 新宿校の生徒のみを抽出
        const shinjukuStudents = properData.filter(row => {
            const campus = row['校舎'];
            return campus && campus.toString().includes('新宿校');
        });

        // 生徒名が空でないもののみを抽出
        const validStudents = shinjukuStudents.filter(student => 
            student['生徒氏名'] && student['生徒氏名'].toString().trim() !== ''
        );

        if (validStudents.length === 0) {
            throw new Error('新宿校の生徒が見つかりません。「校舎」列の内容を確認してください。');
        }

        // テンプレート生成
        generateTemplates(validStudents);
        loading.style.display = 'none';

    } catch (err) {
        console.error('Error processing file:', err);
        showError(`ファイル処理エラー: ${err.message}`);
    }
}

// ファイル読み込み
function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(new Uint8Array(e.target.result));
        reader.onerror = () => reject(new Error('ファイルの読み込みに失敗しました'));
        reader.readAsArrayBuffer(file);
    });
}

// テンプレート生成
function generateTemplates(students) {
    templatesContainer.innerHTML = '';
    
    // 成功メッセージと進捗表示
    const headerDiv = document.createElement('div');
    headerDiv.innerHTML = `
        <div class="success">
            ✅ ${students.length}名の新宿校生徒のテンプレートを生成しました
        </div>
        <div class="progress-summary" id="progressSummary">
            <div class="progress-item">
                📋 コピー済み: <span id="copiedCount">0</span>/${students.length}
            </div>
            <div class="progress-item">
                📱 LINE送信済み: <span id="sentCount">0</span>/${students.length}
            </div>
        </div>
    `;
    templatesContainer.appendChild(headerDiv);

    // 各生徒のテンプレートを生成
    students.forEach((student, index) => {
        const templateDiv = document.createElement('div');
        templateDiv.className = 'template-item';
        
        const template = getTemplate(student);
        const studentName = student['生徒氏名'] || '名前なし';
        const studentId = `student_${index}`;
        
        templateDiv.innerHTML = `
            <div class="template-header">
                <div class="student-info">
                    <div class="student-name">${index + 1}人目: ${escapeHtml(studentName)}</div>
                    <div class="student-status">
                        <span class="copy-status" id="copyStatus_${index}">📝 未コピー</span>
                        <span class="line-status" id="lineStatus_${index}">⏳ LINE未送信</span>
                    </div>
                </div>
                <div class="action-buttons">
                    <button class="copy-btn" data-index="${index}" data-template="${escapeAttribute(template)}">
                        📋 コピー
                    </button>
                    <label class="line-checkbox-label">
                        <input type="checkbox" class="line-checkbox" data-index="${index}" data-student="${escapeAttribute(studentName)}">
                        <span class="checkbox-text">📱 LINE送信完了</span>
                    </label>
                </div>
            </div>
            
            <div class="template-content">${escapeHtml(studentName)}
${escapeHtml(template)}</div>
        `;
        
        templatesContainer.appendChild(templateDiv);
    });

    // イベントリスナーを追加
    addEventListeners();
    
    // 進捗を更新
    updateProgress();

    templatesContainer.style.display = 'block';
}

// イベントリスナーを追加
function addEventListeners() {
    // コピーボタンのイベント
    const copyButtons = templatesContainer.querySelectorAll('.copy-btn');
    copyButtons.forEach(button => {
        button.addEventListener('click', function() {
            const template = this.getAttribute('data-template');
            const index = parseInt(this.getAttribute('data-index'));
            copyToClipboard(template, index, this);
        });
    });

    // チェックボックスのイベント
    const checkboxes = templatesContainer.querySelectorAll('.line-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.addEventListener('change', function() {
            const index = parseInt(this.getAttribute('data-index'));
            const studentName = this.getAttribute('data-student');
            handleLineStatusChange(index, studentName, this.checked);
        });
    });
}

// LINE送信状態変更の処理
function handleLineStatusChange(index, studentName, isChecked) {
    const lineStatusElement = document.getElementById(`lineStatus_${index}`);
    
    if (isChecked) {
        lineSentStatus[index] = {
            studentName: studentName,
            sentAt: new Date().toLocaleString('ja-JP')
        };
        lineStatusElement.textContent = '✅ LINE送信済み';
        lineStatusElement.className = 'line-status sent';
    } else {
        delete lineSentStatus[index];
        lineStatusElement.textContent = '⏳ LINE未送信';
        lineStatusElement.className = 'line-status pending';
    }
    
    updateProgress();
}

// 進捗更新
function updateProgress() {
    const copiedCount = document.querySelectorAll('.copy-status.copied').length;
    const sentCount = Object.keys(lineSentStatus).length;
    
    document.getElementById('copiedCount').textContent = copiedCount;
    document.getElementById('sentCount').textContent = sentCount;
}

// テンプレート文字列生成
function getTemplate(student) {
    const regStart = student['reg開始'] || '';
    const regEnd = student['指導終了'] || '';
    const subject = student['科目'] || '';
    const teacher = student['担当'] || '';

    return `明日の特訓の詳細です。
${regStart}‐${regEnd}
教科：${subject}
担当：${teacher}
お待ちしております。
この通知に返信不要です。`;
}

// HTMLエスケープ
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 属性値エスケープ
function escapeAttribute(str) {
    return str.replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\n/g, '&#10;');
}

// クリップボードにコピー
async function copyToClipboard(text, index, buttonElement) {
    try {
        await navigator.clipboard.writeText(text);
        
        // ボタンの表示を変更
        const originalText = buttonElement.innerHTML;
        buttonElement.innerHTML = '✅ コピー済み';
        buttonElement.classList.add('copied');
        
        // コピー状態を更新
        const copyStatusElement = document.getElementById(`copyStatus_${index}`);
        copyStatusElement.textContent = '✅ コピー済み';
        copyStatusElement.className = 'copy-status copied';
        
        // 進捗を更新
        updateProgress();
        
        // 2秒後に元に戻す
        setTimeout(() => {
            buttonElement.innerHTML = originalText;
            buttonElement.classList.remove('copied');
        }, 2000);
        
    } catch (err) {
        console.error('コピーに失敗しました:', err);
        
        // フォールバック: テキストエリアを使用
        try {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
            
            // 成功時の表示変更
            const originalText = buttonElement.innerHTML;
            buttonElement.innerHTML = '✅ コピー済み';
            buttonElement.classList.add('copied');
            
            // コピー状態を更新
            const copyStatusElement = document.getElementById(`copyStatus_${index}`);
            copyStatusElement.textContent = '✅ コピー済み';
            copyStatusElement.className = 'copy-status copied';
            
            updateProgress();
            
            setTimeout(() => {
                buttonElement.innerHTML = originalText;
                buttonElement.classList.remove('copied');
            }, 2000);
            
        } catch (fallbackErr) {
            console.error('フォールバックコピーも失敗:', fallbackErr);
            alert('コピーに失敗しました。手動でコピーしてください。');
        }
    }
}

// デバッグ用: コンソールに情報を出力
function debugLog(message, data = null) {
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
        console.log('Debug:', message, data);
    }
}
