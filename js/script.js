// ماژول مدیریت رابط کاربری
const UiModule = (function () {
    const elements = {
        dropzone: document.getElementById('dropzone'),
        fileInput: document.getElementById('fileInput'),
        fileInputLabel: document.getElementById('fileInputLabel'),
        namesList: document.getElementById('namesList'),
        confirmBtn: document.getElementById('confirmBtn'),
        clearBtn: document.getElementById('clearBtn'),
        startDrawBtn: document.getElementById('startDrawBtn'),
        redoBtn: document.getElementById('redoBtn'),
        clearResultsBtn: document.getElementById('clearResultsBtn'),
        wheel: document.getElementById('wheel'),
        pointer: document.getElementById('pointer'),
        results: document.getElementById('results'),
        loading: document.getElementById('loading'),
        searchInput: document.getElementById('searchInput'),
        totalNames: document.getElementById('totalNames'),
        winnersCount: document.getElementById('winnersCount'),
        remainingNames: document.getElementById('remainingNames'),
        drawCount: document.getElementById('drawCount'),
        exampleBtn: document.getElementById('exampleBtn'),
        filePreview: document.getElementById('filePreview'),
        filePreviewList: document.getElementById('filePreviewList'),
        progressFill: document.getElementById('progressFill'),
        lastDrawTime: document.getElementById('lastDrawTime'),
        systemStatus: document.getElementById('systemStatus'),
        winnersCountInput: document.getElementById('winnersCount'),
        autoDrawInput: document.getElementById('autoDraw'),
        quickDrawBtn: document.getElementById('quickDrawBtn'),
        exportExcelBtn: document.getElementById('exportExcelBtn'),
        copyResultsBtn: document.getElementById('copyResultsBtn'),
        winnersChart: document.getElementById('winnersChart')
    };

    // نمایش پیام
    function toast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.textContent = message;
        toast.setAttribute('role', 'alert');
        toast.setAttribute('aria-live', 'assertive');

        if (type === 'success') {
            toast.style.background = 'linear-gradient(90deg, #6ee7b7, #34d399)';
        } else if (type === 'error') {
            toast.style.background = 'linear-gradient(90deg, #ff7b7b, #f87171)';
        } else if (type === 'warning') {
            toast.style.background = 'linear-gradient(90deg, #ffd166, #fbbf24)';
        }

        document.getElementById('toastHolder').appendChild(toast);

        setTimeout(() => {
            toast.classList.add('show');
        }, 10);

        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => {
                toast.remove();
            }, 300);
        }, 3500);
    }

    // نمایش لودینگ
    function showLoading() {
        elements.loading.classList.add('active');
    }

    function hideLoading() {
        elements.loading.classList.remove('active');
    }

    // غیرفعال کردن بخش آپلود
    function disableUploadSection() {
        elements.dropzone.classList.add('disabled');
        elements.fileInputLabel.style.pointerEvents = 'none';
        elements.fileInput.disabled = true;
    }

    // فعال کردن بخش آپلود
    function enableUploadSection() {
        elements.dropzone.classList.remove('disabled');
        elements.fileInputLabel.style.pointerEvents = 'auto';
        elements.fileInput.disabled = false;
    }

    // نمایش پیش‌نمایش فایل
    function showFilePreview(names) {
        if (!names || names.length === 0) {
            elements.filePreview.style.display = 'none';
            return;
        }

        elements.filePreviewList.innerHTML = '';
        const previewCount = Math.min(names.length, 10);

        for (let i = 0; i < previewCount; i++) {
            const item = document.createElement('div');
            item.className = 'file-preview-item';
            item.textContent = `${i + 1}. ${names[i]}`;
            elements.filePreviewList.appendChild(item);
        }

        if (names.length > 10) {
            const moreItem = document.createElement('div');
            moreItem.className = 'file-preview-item';
            moreItem.textContent = `... و ${names.length - 10} نام دیگر`;
            elements.filePreviewList.appendChild(moreItem);
        }

        elements.filePreview.style.display = 'block';
    }

    // نمایش اسامی قابل ویرایش
    function renderNamesEditable(names) {
        elements.namesList.innerHTML = '';

        if (!names || names.length === 0) {
            elements.namesList.innerHTML = `
            <div class="empty-state">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
                <path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
              </svg>
              <p>هیچ اسمی پیدا نشد</p>
            </div>
          `;
            elements.confirmBtn.disabled = true;
            elements.clearBtn.disabled = true;
            return;
        }

        names.forEach((name, index) => {
            const row = document.createElement('div');
            row.className = 'name-item';

            const number = document.createElement('div');
            number.className = 'number';
            number.textContent = index + 1;

            const nameInput = document.createElement('input');
            nameInput.className = 'name';
            nameInput.value = name;
            nameInput.setAttribute('data-idx', index);

            const actions = document.createElement('div');
            actions.className = 'name-actions';

            const deleteBtn = document.createElement('button');
            deleteBtn.className = 'action-btn';
            deleteBtn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
          `;
            deleteBtn.addEventListener('click', () => {
                row.remove();
                updateConfirmButtonState();
            });

            actions.appendChild(deleteBtn);
            row.appendChild(number);
            row.appendChild(nameInput);
            row.appendChild(actions);
            elements.namesList.appendChild(row);
        });

        elements.confirmBtn.disabled = false;
        elements.clearBtn.disabled = false;
    }

    // دریافت اسامی از رابط کاربری
    function getNamesFromUi() {
        const inputs = elements.namesList.querySelectorAll('input');
        const names = [];
        inputs.forEach(input => {
            const value = input.value.trim();
            if (value) names.push(value);
        });
        return names;
    }

    // به‌روزرسانی وضعیت دکمه تأیید
    function updateConfirmButtonState() {
        const hasNames = elements.namesList.querySelectorAll('input').length > 0;
        elements.confirmBtn.disabled = !hasNames;
        elements.clearBtn.disabled = !hasNames;
    }

    // نمایش اسامی در چرخ قرعه‌کشی
    function renderWheelItems(names) {
        elements.wheel.innerHTML = '';

        if (!names || names.length === 0) {
            elements.wheel.innerHTML = `
            <div class="empty-state">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="8" x2="12" y2="12"></line>
                <line x1="12" y1="16" x2="12.01" y2="16"></line>
              </svg>
              <p>اسامی برای قرعه‌کشی نمایش داده می‌شود</p>
            </div>
          `;
            return;
        }

        names.forEach((name, index) => {
            const item = document.createElement('div');
            item.className = 'wheel-item';
            item.id = `wheel-item-${index}`;
            item.innerHTML = `
            <div class="number">${index + 1}</div>
            <div class="name">${name}</div>
          `;
            elements.wheel.appendChild(item);
        });
    }

    // نمایش نتایج
    function renderResults(results) {
        elements.results.innerHTML = '';

        if (!results || results.length === 0) {
            elements.results.innerHTML = `
            <div class="empty-state">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
                <path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <p>نتیجه‌ای هنوز تولید نشده است</p>
            </div>
          `;
            elements.clearResultsBtn.disabled = true;
            elements.exportExcelBtn.disabled = true;
            elements.copyResultsBtn.disabled = true;
            return;
        }

        results.forEach((result, index) => {
            const item = document.createElement('div');
            item.className = 'result-item';
            item.innerHTML = `
            <div class="rank">#${index + 1}</div>
            <div class="name">${result}</div>
          `;
            elements.results.appendChild(item);
        });

        elements.clearResultsBtn.disabled = false;
        elements.exportExcelBtn.disabled = false;
        elements.copyResultsBtn.disabled = false;
    }

    // به‌روزرسانی آمار
    function updateStats(total, winners, remaining, draws) {
        elements.totalNames.textContent = total;
        elements.winnersCount.textContent = winners;
        elements.remainingNames.textContent = remaining;
        elements.drawCount.textContent = draws;
    }

    // به‌روزرسانی وضعیت سیستم
    function updateSystemStatus(status, progress = 0) {
        elements.systemStatus.textContent = status;
        elements.progressFill.style.width = `${progress}%`;

        if (progress > 0) {
            elements.progressFill.style.display = 'block';
        } else {
            elements.progressFill.style.display = 'none';
        }
    }

    // به‌روزرسانی زمان آخرین قرعه‌کشی
    function updateLastDrawTime() {
        const now = new Date();
        elements.lastDrawTime.textContent = now.toLocaleTimeString('fa-IR');
    }

    // نمایش نمودار
    function renderWinnersChart(results) {
        elements.winnersChart.innerHTML = '';

        if (!results || results.length === 0) {
            elements.winnersChart.innerHTML = `
                        <div class="empty-state">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
                                <path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                            </svg>
                            <p>داده‌ای برای نمایش وجود ندارد</p>
                        </div>
                    `;
            return;
        }

        // گروه‌بندی نتایج
        const groups = {};
        results.forEach(result => {
            groups[result] = (groups[result] || 0) + 1;
        });

        const maxCount = Math.max(...Object.values(groups));
        const chartHeight = 150;

        Object.entries(groups).forEach(([name, count]) => {
            const bar = document.createElement('div');
            bar.className = 'chart-bar';
            const height = (count / maxCount) * chartHeight;
            bar.style.height = `${height}px`;

            const label = document.createElement('div');
            label.className = 'chart-label';
            label.textContent = name.split(' ')[0]; // فقط نام کوچک

            bar.appendChild(label);
            elements.winnersChart.appendChild(bar);
        });
    }

    return {
        elements,
        toast,
        showLoading,
        hideLoading,
        disableUploadSection,
        enableUploadSection,
        showFilePreview,
        renderNamesEditable,
        getNamesFromUi,
        updateConfirmButtonState,
        renderWheelItems,
        renderResults,
        updateStats,
        updateSystemStatus,
        updateLastDrawTime,
        renderWinnersChart
    };
})();

// ماژول مدیریت فایل اکسل
const ExcelModule = (function () {
    // خواندن فایل اکسل
    function readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = (error) => reject(error);
            reader.readAsArrayBuffer(file);
        });
    }

    // استخراج اسامی از فایل اکسل - فقط ردیف اول هدر است
    function extractNamesFromWorkbook(workbook) {
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) return [];

        const sheet = workbook.Sheets[firstSheetName];
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
        const names = [];

        // خواندن از ردیف دوم به بعد (ردیف 1 به بعد در index) - ردیف اول هدر است
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            // ستون B (شماره 1)
            const cellB = { c: 1, r: row };
            const cellRefB = XLSX.utils.encode_cell(cellB);
            const cellValueB = sheet[cellRefB];

            // ستون C (شماره 2)
            const cellC = { c: 2, r: row };
            const cellRefC = XLSX.utils.encode_cell(cellC);
            const cellValueC = sheet[cellRefC];

            let valueB = '';
            let valueC = '';

            if (cellValueB && cellValueB.v !== undefined) {
                valueB = String(cellValueB.v).trim();
            }

            if (cellValueC && cellValueC.v !== undefined) {
                valueC = String(cellValueC.v).trim();
            }

            // ترکیب مقادیر دو ستون
            const combined = valueB && valueC ? `${valueB} ${valueC}` : valueB || valueC;
            if (combined) names.push(combined);
        }

        return names;
    }

    // تولید فایل اکسل از نتایج
    function generateExcelFromResults(results) {
        const worksheet = XLSX.utils.aoa_to_sheet([
            ['ردیف', 'نام برنده'],
            ...results.map((result, index) => [index + 1, result])
        ]);

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'نتایج قرعه‌کشی');

        return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    }

    return { readFile, extractNamesFromWorkbook, generateExcelFromResults };
})();

// ماژول قرعه‌کشی
const DrawModule = (function () {
    // درهم‌سازی آرایه
    function shuffle(array) {
        const newArray = [...array];
        for (let i = newArray.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
        }
        return newArray;
    }

    // انیمیشن قرعه‌کشی
    function animateDraw(names, onStep, onComplete) {
        if (!names || names.length === 0) {
            onComplete(null);
            return () => { };
        }

        const shuffled = shuffle(names);
        let currentIndex = 0;
        let ticks = 0;
        let speed = 80;
        const maxTicks = 50 + Math.floor(Math.random() * 20);

        const intervalId = setInterval(() => {
            // حذف انتخاب قبلی
            document.querySelectorAll('.wheel-item').forEach(item => {
                item.classList.remove('selected');
            });

            // انتخاب آیتم فعلی
            const currentItem = document.getElementById(`wheel-item-${currentIndex}`);
            if (currentItem) {
                currentItem.classList.add('selected');

                // به‌روزرسانی موقعیت اشاره‌گر
                const itemRect = currentItem.getBoundingClientRect();
                const wheelRect = document.getElementById('wheel').getBoundingClientRect();
                const pointer = document.getElementById('pointer');

                pointer.style.top = `${itemRect.top - wheelRect.top + itemRect.height / 2}px`;
            }

            // فراخوانی تابع مرحله
            onStep(shuffled[currentIndex], currentIndex);

            // به‌روزرسانی شاخص
            currentIndex = (currentIndex + 1) % shuffled.length;
            ticks++;

            // کاهش تدریجی سرعت
            if (ticks > maxTicks * 0.7) speed += 30;
            else if (ticks > maxTicks * 0.4) speed += 15;

            // پایان انیمیشن
            if (ticks >= maxTicks) {
                clearInterval(intervalId);

                // انتخاب نهایی
                const finalIndex = Math.floor(Math.random() * shuffled.length);
                const finalName = shuffled[finalIndex];

                // نمایش آیتم انتخاب شده نهایی
                document.querySelectorAll('.wheel-item').forEach(item => {
                    item.classList.remove('selected');
                });

                const finalItem = document.getElementById(`wheel-item-${finalIndex}`);
                if (finalItem) {
                    finalItem.classList.add('selected');
                    finalItem.classList.add('winner-glow');

                    // به‌روزرسانی موقعیت اشاره‌گر
                    const itemRect = finalItem.getBoundingClientRect();
                    const wheelRect = document.getElementById('wheel').getBoundingClientRect();
                    const pointer = document.getElementById('pointer');

                    pointer.style.top = `${itemRect.top - wheelRect.top + itemRect.height / 2}px`;
                }

                onComplete(finalName);
            }
        }, speed);

        return () => clearInterval(intervalId);
    }

    // قرعه‌کشی چند نفره
    function drawMultiple(names, count) {
        if (!names || names.length === 0 || count <= 0) return [];

        const shuffled = shuffle(names);
        return shuffled.slice(0, Math.min(count, names.length));
    }

    return { shuffle, animateDraw, drawMultiple };
})();

// ایجاد ذرات متحرک
function createParticles() {
    const particlesContainer = document.getElementById('particles');
    const colors = ['#5b8cff', '#6ee7b7', '#ffd166', '#ff7b7b'];

    for (let i = 0; i < 30; i++) {
        const particle = document.createElement('div');
        particle.className = 'particle';

        const size = Math.random() * 10 + 5;
        particle.style.width = `${size}px`;
        particle.style.height = `${size}px`;

        particle.style.left = `${Math.random() * 100}vw`;
        particle.style.top = `${Math.random() * 100}vh`;

        particle.style.background = colors[Math.floor(Math.random() * colors.length)];
        particle.style.animationDelay = `${Math.random() * 15}s`;
        particle.style.animationDuration = `${15 + Math.random() * 10}s`;

        particlesContainer.appendChild(particle);
    }
}

// ایجاد کانفی
function createConfetti() {
    const colors = ['#5b8cff', '#6ee7b7', '#ffd166', '#ff7b7b', '#a78bfa', '#f472b6'];

    for (let i = 0; i < 100; i++) {
        const confetti = document.createElement('div');
        confetti.className = 'confetti';

        confetti.style.left = `${Math.random() * 100}vw`;
        confetti.style.background = colors[Math.floor(Math.random() * colors.length)];

        const size = Math.random() * 10 + 5;
        confetti.style.width = `${size}px`;
        confetti.style.height = `${size}px`;

        if (Math.random() > 0.5) {
            confetti.style.borderRadius = '50%';
        }

        confetti.style.animation = `confetti-fall ${2 + Math.random() * 3}s linear forwards`;
        confetti.style.animationDelay = `${Math.random() * 2}s`;

        document.body.appendChild(confetti);

        setTimeout(() => {
            confetti.remove();
        }, 5000);
    }
}

// تغییر تم
function setTheme(themeName) {
    document.body.className = `${themeName}-theme`;

    document.querySelectorAll('.theme-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.querySelector(`.theme-${themeName}`).classList.add('active');

    localStorage.setItem('selectedTheme', themeName);
}

// برنامه اصلی
const App = (function (Ui, Excel, Draw) {
    const e = Ui.elements;
    let currentNames = [];
    let confirmedNames = [];
    let results = [];
    let cancelAnimation = null;
    let drawCount = 0;
    let isConfirmed = false;
    let autoDrawInterval = null;

    function init() {
        setupEventListeners();
        createParticles();

        // بارگذاری تم ذخیره شده
        const savedTheme = localStorage.getItem('selectedTheme') || 'light';
        setTheme(savedTheme);

        // رویدادهای تغییر تم
        document.querySelectorAll('.theme-btn').forEach(btn => {
            btn.addEventListener('click', function () {
                setTheme(this.getAttribute('data-theme'));
            });
        });

        // راهنمای صفحه
        showHelpOnFirstVisit();
    }

    function setupEventListeners() {
        // درگ و دراپ
        ['dragenter', 'dragover'].forEach(event => {
            e.dropzone.addEventListener(event, (e) => {
                if (!isConfirmed) {
                    e.preventDefault();
                    e.dropzone.classList.add('dragover');
                }
            });
        });

        ['dragleave', 'drop'].forEach(event => {
            e.dropzone.addEventListener(event, (e) => {
                if (!isConfirmed) {
                    e.preventDefault();
                    e.dropzone.classList.remove('dragover');
                }
            });
        });

        e.dropzone.addEventListener('drop', (e) => {
            if (!isConfirmed) {
                const file = e.dataTransfer?.files[0];
                if (file) processFile(file);
            }
        });

        // انتخاب فایل
        e.fileInput.addEventListener('change', (e) => {
            if (!isConfirmed) {
                const file = e.target.files[0];
                if (file) processFile(file);
            }
        });

        // دکمه‌ها
        e.confirmBtn.addEventListener('click', handleConfirm);
        e.clearBtn.addEventListener('click', handleClear);
        e.startDrawBtn.addEventListener('click', handleStartDraw);
        e.redoBtn.addEventListener('click', handleRedo);
        e.clearResultsBtn.addEventListener('click', handleClearResults);
        e.exampleBtn.addEventListener('click', loadExample);
        e.quickDrawBtn.addEventListener('click', handleQuickDraw);
        e.exportExcelBtn.addEventListener('click', handleExportExcel);
        e.copyResultsBtn.addEventListener('click', handleCopyResults);

        // جستجو
        e.searchInput.addEventListener('input', handleSearch);

        // کنترل‌های پیشرفته
        e.autoDrawInput.addEventListener('change', handleAutoDrawChange);

        // دسترسی‌پذیری - کلیدهای میانبر
        document.addEventListener('keydown', handleKeyboardShortcuts);
    }

    async function processFile(file) {
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            Ui.toast('فایل انتخاب شده با فرمت .xlsx نیست.', 'error');
            return;
        }

        Ui.showLoading();
        Ui.updateSystemStatus('در حال پردازش فایل...', 30);
        try {
            const workbook = await Excel.readFile(file);
            Ui.updateSystemStatus('در حال استخراج اسامی...', 60);
            const names = Excel.extractNamesFromWorkbook(workbook);

            if (!names || names.length === 0) {
                Ui.renderNamesEditable([]);
                Ui.showFilePreview([]);
                Ui.toast('هیچ اسمی در ستون‌های B و C یافت نشد.', 'warning');
                return;
            }

            currentNames = names;
            Ui.renderNamesEditable(names);
            Ui.showFilePreview(names);
            Ui.updateSystemStatus('آماده', 100);
            Ui.toast(`تعداد ${names.length} اسم با موفقیت بارگذاری شد.`, 'success');

            // به‌روزرسانی نمودار
            setTimeout(() => Ui.renderWinnersChart(results), 100);
        } catch (error) {
            console.error(error);
            Ui.toast('خطا در خواندن فایل اکسل.', 'error');
            Ui.updateSystemStatus('خطا', 0);
        } finally {
            Ui.hideLoading();
        }
    }

    function handleConfirm() {
        const names = Ui.getNamesFromUi();
        if (!names || names.length === 0) {
            Ui.toast('لیست اسامی خالی است.', 'warning');
            return;
        }

        confirmedNames = names;
        isConfirmed = true;

        Ui.renderWheelItems(names);
        Ui.disableUploadSection();
        e.startDrawBtn.disabled = false;
        e.quickDrawBtn.disabled = false;
        updateStats();
        Ui.toast('اسامی تأیید شد. بخش آپلود غیرفعال شد.', 'success');
    }

    function handleClear() {
        currentNames = [];
        confirmedNames = [];
        results = [];
        drawCount = 0;
        isConfirmed = false;

        Ui.renderNamesEditable([]);
        Ui.renderWheelItems([]);
        Ui.renderResults([]);
        Ui.showFilePreview([]);
        Ui.enableUploadSection();
        e.startDrawBtn.disabled = true;
        e.redoBtn.disabled = true;
        e.quickDrawBtn.disabled = true;
        updateStats();

        // توقف قرعه‌کشی خودکار
        if (autoDrawInterval) {
            clearInterval(autoDrawInterval);
            autoDrawInterval = null;
        }

        Ui.toast('همه اسامی و نتایج پاک شد. بخش آپلود فعال شد.', 'info');
    }

    function handleClearResults() {
        results = [];
        Ui.renderResults([]);
        updateStats();
        Ui.toast('نتایج پاک شد.', 'info');
    }

    function handleStartDraw() {
        if (!confirmedNames || confirmedNames.length === 0) {
            Ui.toast('هیچ نامی برای قرعه‌کشی موجود نیست.', 'warning');
            return;
        }

        e.startDrawBtn.disabled = true;
        e.redoBtn.disabled = true;
        e.quickDrawBtn.disabled = true;

        const winnersCount = parseInt(e.winnersCountInput.value) || 1;

        if (winnersCount > 1) {
            handleMultipleDraw(winnersCount);
            return;
        }

        cancelAnimation = Draw.animateDraw(
            confirmedNames,
            (name, index) => {
                Ui.updateSystemStatus('در حال قرعه‌کشی...', 50);
            },
            (winner) => {
                if (winner) {
                    results.push(winner);
                    drawCount++;

                    // حذف برنده از لیست اسامی تأیید شده
                    const winnerIndex = confirmedNames.indexOf(winner);
                    if (winnerIndex !== -1) {
                        confirmedNames.splice(winnerIndex, 1);
                        Ui.renderWheelItems(confirmedNames);
                    }

                    Ui.renderResults(results);
                    Ui.renderWinnersChart(results);
                    createConfetti();
                    Ui.updateLastDrawTime();
                    Ui.toast(`برنده: ${winner}`, 'success');
                }

                e.startDrawBtn.disabled = confirmedNames.length === 0;
                e.redoBtn.disabled = false;
                e.quickDrawBtn.disabled = false;
                updateStats();
                Ui.updateSystemStatus('آماده', 100);
            }
        );
    }

    function handleMultipleDraw(count) {
        if (confirmedNames.length < count) {
            Ui.toast(`تعداد اسامی باقیمانده (${confirmedNames.length}) کمتر از تعداد درخواستی (${count}) است.`, 'warning');
            e.startDrawBtn.disabled = false;
            e.redoBtn.disabled = false;
            e.quickDrawBtn.disabled = false;
            return;
        }

        const winners = Draw.drawMultiple(confirmedNames, count);

        winners.forEach(winner => {
            results.push(winner);
            const winnerIndex = confirmedNames.indexOf(winner);
            if (winnerIndex !== -1) {
                confirmedNames.splice(winnerIndex, 1);
            }
        });

        drawCount++;
        Ui.renderWheelItems(confirmedNames);
        Ui.renderResults(results);
        Ui.renderWinnersChart(results);
        createConfetti();
        Ui.updateLastDrawTime();

        if (winners.length === 1) {
            Ui.toast(`برنده: ${winners[0]}`, 'success');
        } else {
            Ui.toast(`${winners.length} برنده انتخاب شدند.`, 'success');
        }

        e.startDrawBtn.disabled = confirmedNames.length === 0;
        e.redoBtn.disabled = false;
        e.quickDrawBtn.disabled = false;
        updateStats();
    }

    function handleQuickDraw() {
        if (!confirmedNames || confirmedNames.length === 0) {
            Ui.toast('هیچ نامی برای قرعه‌کشی موجود نیست.', 'warning');
            return;
        }

        const winnersCount = parseInt(e.winnersCountInput.value) || 1;
        const winners = Draw.drawMultiple(confirmedNames, winnersCount);

        winners.forEach(winner => {
            results.push(winner);
            const winnerIndex = confirmedNames.indexOf(winner);
            if (winnerIndex !== -1) {
                confirmedNames.splice(winnerIndex, 1);
            }
        });

        drawCount++;
        Ui.renderWheelItems(confirmedNames);
        Ui.renderResults(results);
        Ui.renderWinnersChart(results);
        Ui.updateLastDrawTime();

        if (winners.length === 1) {
            Ui.toast(`برنده: ${winners[0]}`, 'success');
        } else {
            Ui.toast(`${winners.length} برنده انتخاب شدند.`, 'success');
        }

        e.startDrawBtn.disabled = confirmedNames.length === 0;
        updateStats();
    }

    function handleRedo() {
        e.startDrawBtn.disabled = confirmedNames.length === 0;
        e.redoBtn.disabled = true;
    }

    function handleSearch() {
        const searchTerm = e.searchInput.value.toLowerCase();
        const nameItems = e.namesList.querySelectorAll('.name-item');

        nameItems.forEach(item => {
            const nameInput = item.querySelector('input');
            const name = nameInput.value.toLowerCase();
            item.style.display = name.includes(searchTerm) ? 'flex' : 'none';
        });
    }

    function handleAutoDrawChange() {
        const interval = parseInt(e.autoDrawInput.value);

        if (autoDrawInterval) {
            clearInterval(autoDrawInterval);
            autoDrawInterval = null;
        }

        if (interval > 0 && confirmedNames.length > 0) {
            autoDrawInterval = setInterval(() => {
                if (confirmedNames.length > 0) {
                    handleQuickDraw();
                } else {
                    clearInterval(autoDrawInterval);
                    autoDrawInterval = null;
                    Ui.toast('قرعه‌کشی خودکار به پایان رسید.', 'info');
                }
            }, interval * 1000);

            Ui.toast(`قرعه‌کشی خودکار هر ${interval} ثانیه فعال شد.`, 'success');
        }
    }

    function handleExportExcel() {
        try {
            const excelData = Excel.generateExcelFromResults(results);
            const blob = new Blob([excelData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'نتایج_قرعه_کشی.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            Ui.toast('فایل Excel با موفقیت ذخیره شد.', 'success');
        } catch (error) {
            Ui.toast('خطا در تولید فایل Excel.', 'error');
        }
    }

    function handleCopyResults() {
        const resultsText = results.map((result, index) => `${index + 1}. ${result}`).join('\n');
        navigator.clipboard.writeText(resultsText).then(() => {
            Ui.toast('نتایج با موفقیت کپی شد.', 'success');
        }).catch(() => {
            // Fallback برای مرورگرهای قدیمی
            const textArea = document.createElement('textarea');
            textArea.value = resultsText;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            Ui.toast('نتایج با موفقیت کپی شد.', 'success');
        });
    }

    function loadExample() {
        if (isConfirmed) {
            Ui.toast('برای بارگذاری مثال جدید، ابتدا اسامی فعلی را پاک کنید.', 'warning');
            return;
        }

        const exampleNames = [
            'علی رضایی',
            'سارا محمدی',
            'مهدی کریمی',
            'نرگس حسینی',
            'پوریا ملکی',
            'فاطمه احمدی',
            'محمد جعفری',
            'زهرا امیری',
            'حسین قربانی',
            'لیلا موسوی'
        ];

        currentNames = exampleNames;
        Ui.renderNamesEditable(exampleNames);
        Ui.showFilePreview(exampleNames);
        Ui.toast('مثال نمونه بارگذاری شد.', 'success');
    }

    function handleKeyboardShortcuts(e) {
        // فقط وقتی که focus روی المنت خاصی نیست
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

        if (e.key === 'Enter' && !e.startDrawBtn.disabled) {
            handleStartDraw();
        } else if (e.key === 'Escape') {
            handleClear();
        } else if (e.key === 'r' || e.key === 'R') {
            handleRedo();
        }
    }

    function showHelpOnFirstVisit() {
        const hasVisited = localStorage.getItem('hasVisited');
        if (!hasVisited) {
            setTimeout(() => {
                Ui.toast('برای راهنمایی، نشانگر موس را روی آیکون‌های راهنما نگه دارید.', 'info');
                localStorage.setItem('hasVisited', 'true');
            }, 2000);
        }
    }

    function updateStats() {
        Ui.updateStats(
            currentNames.length,
            results.length,
            confirmedNames.length,
            drawCount
        );
    }

    return { init };
})(UiModule, ExcelModule, DrawModule);

// آغاز برنامه
document.addEventListener('DOMContentLoaded', function () {
    App.init();
});