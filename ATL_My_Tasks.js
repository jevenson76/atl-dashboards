// ATL My Tasks - Interactive Task Management
// Safe DOM manipulation for SharePoint embedding

(function() {
    'use strict';

    // Sample task data - PMO Lead (Jason Evenson) tasks for demo
    var myTasks = [
        {
            id: 'T003',
            title: 'Reporting Standards & Dashboard Automation',
            phase: 'Phase 1: Stabilization',
            workstream: 'PMO Foundation',
            status: 'completed',
            progress: 100,
            dueDate: '2025-12-19',
            owner: 'Jason Evenson',
            notes: [
                { date: '2025-12-19', text: 'Dashboard deployed to SharePoint. KPIs updating hourly via Power Automate.' },
                { date: '2025-12-15', text: 'Reporting standards document finalized and approved by leadership.' }
            ],
            attachments: ['Reporting_Standards_v2.pdf']
        },
        {
            id: 'T006',
            title: 'SKU Remediation - Cataloging & Analysis',
            phase: 'Phase 1: Stabilization',
            workstream: 'SKU Remediation',
            status: 'in-progress',
            progress: 45,
            dueDate: '2026-01-16',
            owner: 'Jason Evenson',
            notes: [
                { date: '2026-01-06', text: 'Identified 450 SKUs requiring review. Working with ATL ops on prioritization.' },
                { date: '2025-12-28', text: 'Initial SKU export received from ATL systems.' }
            ],
            attachments: ['SKU_Analysis_Draft.xlsx']
        },
        {
            id: 'T011',
            title: 'Training Program Development',
            phase: 'Phase 2: Foundation',
            workstream: 'Change Management',
            status: 'in-progress',
            progress: 30,
            dueDate: '2026-02-06',
            owner: 'Jason Evenson',
            notes: [
                { date: '2026-01-05', text: 'Outlined training modules for SharePoint, task updates, and reporting.' }
            ],
            attachments: []
        },
        {
            id: 'T012',
            title: 'CRM/System Integration Planning',
            phase: 'Phase 2: Foundation',
            workstream: 'Systems',
            status: 'at-risk',
            progress: 20,
            dueDate: '2026-02-06',
            owner: 'Jason Evenson',
            notes: [
                { date: '2026-01-07', text: 'Need IT resource assignment. Escalating to Jennifer.' }
            ],
            attachments: []
        },
        {
            id: 'T015',
            title: 'Weekly PMO Cadence - First Sync Setup',
            phase: 'Phase 1: Stabilization',
            workstream: 'PMO Operations',
            status: 'in-progress',
            progress: 75,
            dueDate: '2026-01-15',
            owner: 'Jason Evenson',
            notes: [
                { date: '2026-01-08', text: 'Calendar invites sent. Agenda template created.' }
            ],
            attachments: ['Weekly_Sync_Agenda_Template.docx']
        },
        {
            id: 'T016',
            title: 'Planner Board Configuration',
            phase: 'Phase 1: Stabilization',
            workstream: 'PMO Operations',
            status: 'blocked',
            progress: 50,
            dueDate: '2026-01-10',
            owner: 'Jason Evenson',
            notes: [
                { date: '2026-01-07', text: 'Blocked - need Teams group permissions from IT. Ticket submitted.' }
            ],
            attachments: []
        }
    ];

    // State
    var currentFilter = 'all';
    var searchTerm = '';
    var expandedTaskId = null;

    // Initialize the page
    function init() {
        initUser();
        updateStats();
        renderTasks();
        setupFilters();
        setupSearch();
        updateFooterDate();
    }

    // Initialize user display
    function initUser() {
        var userName = 'Jason Evenson';
        var userInitials = userName.split(' ').map(function(n) { return n[0]; }).join('');

        var avatarEl = document.getElementById('userAvatar');
        var nameEl = document.getElementById('userName');

        if (avatarEl) avatarEl.textContent = userInitials;
        if (nameEl) nameEl.textContent = userName;
    }

    // Update statistics display
    function updateStats() {
        var stats = {
            total: myTasks.length,
            onTrack: myTasks.filter(function(t) { return t.status === 'on-track' || (t.status === 'in-progress' && t.progress >= 50); }).length,
            inProgress: myTasks.filter(function(t) { return t.status === 'in-progress'; }).length,
            atRisk: myTasks.filter(function(t) { return t.status === 'at-risk'; }).length,
            blocked: myTasks.filter(function(t) { return t.status === 'blocked'; }).length
        };

        setElementText('statTotal', stats.total);
        setElementText('statOnTrack', stats.onTrack);
        setElementText('statInProgress', stats.inProgress);
        setElementText('statAtRisk', stats.atRisk);
        setElementText('statBlocked', stats.blocked);
    }

    function setElementText(id, text) {
        var el = document.getElementById(id);
        if (el) el.textContent = text;
    }

    function renderTasks() {
        var taskList = document.getElementById('taskList');
        if (!taskList) return;

        while (taskList.firstChild) {
            taskList.removeChild(taskList.firstChild);
        }

        var filteredTasks = myTasks.filter(function(task) {
            var matchesFilter = false;
            if (currentFilter === 'all') {
                matchesFilter = true;
            } else if (currentFilter === 'on-track') {
                // On-track is a computed status: in-progress tasks with >= 50% progress
                matchesFilter = task.status === 'in-progress' && task.progress >= 50;
            } else {
                matchesFilter = task.status === currentFilter;
            }
            var matchesSearch = !searchTerm ||
                task.title.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1 ||
                task.workstream.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1 ||
                task.id.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1;
            return matchesFilter && matchesSearch;
        });

        if (filteredTasks.length === 0) {
            var emptyState = document.createElement('div');
            emptyState.className = 'empty-state';
            var emptyText = document.createElement('p');
            emptyText.textContent = currentFilter === 'all' ?
                'No tasks assigned to you.' :
                'No tasks match the current filter.';
            emptyState.appendChild(emptyText);
            taskList.appendChild(emptyState);
            return;
        }

        filteredTasks.forEach(function(task) {
            var card = createTaskCard(task);
            taskList.appendChild(card);
        });
    }

    function createTaskCard(task) {
        var card = document.createElement('div');
        card.className = 'task-card';
        card.dataset.taskId = task.id;
        if (expandedTaskId === task.id) {
            card.classList.add('expanded');
        }

        var header = createTaskHeader(task);
        card.appendChild(header);

        var details = createTaskDetails(task);
        card.appendChild(details);

        return card;
    }

    function createChevronIcon() {
        var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('viewBox', '0 0 24 24');
        svg.setAttribute('fill', 'none');
        svg.setAttribute('stroke', 'currentColor');
        svg.setAttribute('stroke-width', '2');

        var path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('d', 'm6 9 6 6 6-6');
        svg.appendChild(path);

        return svg;
    }

    function createTaskHeader(task) {
        var header = document.createElement('div');
        header.className = 'task-header';
        header.addEventListener('click', function() { toggleTask(task.id); });

        var indicator = document.createElement('div');
        indicator.className = 'task-status-indicator ' + task.status;
        header.appendChild(indicator);

        var main = document.createElement('div');
        main.className = 'task-main';

        var taskId = document.createElement('div');
        taskId.className = 'task-id';
        taskId.textContent = task.id + ' \u2022 ' + task.workstream;
        main.appendChild(taskId);

        var title = document.createElement('div');
        title.className = 'task-title';
        title.textContent = task.title;
        main.appendChild(title);

        var meta = document.createElement('div');
        meta.className = 'task-meta';

        var phaseSpan = document.createElement('span');
        phaseSpan.textContent = task.phase;
        meta.appendChild(phaseSpan);

        var dueSpan = document.createElement('span');
        dueSpan.textContent = 'Due: ' + formatDate(task.dueDate);
        meta.appendChild(dueSpan);

        var statusSpan = document.createElement('span');
        statusSpan.textContent = formatStatus(task.status);
        meta.appendChild(statusSpan);

        main.appendChild(meta);
        header.appendChild(main);

        var progressContainer = document.createElement('div');
        progressContainer.className = 'task-progress';

        var ring = createProgressRing(task.progress);
        progressContainer.appendChild(ring);

        var progressText = document.createElement('div');
        progressText.className = 'progress-text';
        progressText.textContent = task.progress + '%';
        progressContainer.appendChild(progressText);

        header.appendChild(progressContainer);

        var expandBtn = document.createElement('div');
        expandBtn.className = 'expand-btn';
        expandBtn.appendChild(createChevronIcon());
        header.appendChild(expandBtn);

        return header;
    }

    function createProgressRing(progress) {
        var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('class', 'progress-ring');
        svg.setAttribute('viewBox', '0 0 50 50');

        var bgCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        bgCircle.setAttribute('class', 'bg');
        bgCircle.setAttribute('cx', '25');
        bgCircle.setAttribute('cy', '25');
        bgCircle.setAttribute('r', '20');
        svg.appendChild(bgCircle);

        var progressCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        progressCircle.setAttribute('class', 'progress');
        progressCircle.setAttribute('cx', '25');
        progressCircle.setAttribute('cy', '25');
        progressCircle.setAttribute('r', '20');
        var circumference = 2 * Math.PI * 20;
        var offset = circumference - (progress / 100) * circumference;
        progressCircle.setAttribute('stroke-dasharray', String(circumference));
        progressCircle.setAttribute('stroke-dashoffset', String(offset));
        svg.appendChild(progressCircle);

        return svg;
    }

    function createTaskDetails(task) {
        var details = document.createElement('div');
        details.className = 'task-details';

        var grid = document.createElement('div');
        grid.className = 'details-grid';

        var updateSection = createUpdateSection(task);
        grid.appendChild(updateSection);

        var infoSidebar = createInfoSidebar(task);
        grid.appendChild(infoSidebar);

        details.appendChild(grid);
        return details;
    }

    function createUpdateSection(task) {
        var section = document.createElement('div');
        section.className = 'update-section';

        var title = document.createElement('h4');
        title.textContent = 'Update Task';
        section.appendChild(title);

        var statusDiv = document.createElement('div');
        statusDiv.className = 'status-buttons';

        var statuses = [
            { value: 'not-started', label: 'Not Started' },
            { value: 'in-progress', label: 'In Progress' },
            { value: 'completed', label: 'Completed' },
            { value: 'blocked', label: 'Blocked' }
        ];

        statuses.forEach(function(s) {
            var btn = document.createElement('button');
            btn.className = 'status-btn ' + s.value;
            if (task.status === s.value) btn.classList.add('active');
            btn.textContent = s.label;
            btn.dataset.status = s.value;
            btn.dataset.taskId = task.id;
            btn.addEventListener('click', handleStatusClick);
            statusDiv.appendChild(btn);
        });

        section.appendChild(statusDiv);

        var progressControl = document.createElement('div');
        progressControl.className = 'progress-control';

        var progressHeader = document.createElement('div');
        progressHeader.className = 'progress-header';

        var progressLabel = document.createElement('span');
        progressLabel.className = 'progress-label';
        progressLabel.textContent = 'Progress';
        progressHeader.appendChild(progressLabel);

        var progressValue = document.createElement('span');
        progressValue.className = 'progress-value-display';
        progressValue.id = 'progress-display-' + task.id;
        progressValue.textContent = task.progress + '%';
        progressHeader.appendChild(progressValue);

        progressControl.appendChild(progressHeader);

        var slider = document.createElement('input');
        slider.type = 'range';
        slider.className = 'progress-slider';
        slider.id = 'progress-' + task.id;
        slider.min = '0';
        slider.max = '100';
        slider.value = String(task.progress);
        slider.addEventListener('input', handleProgressChange);
        progressControl.appendChild(slider);

        section.appendChild(progressControl);

        var notesTextarea = document.createElement('textarea');
        notesTextarea.className = 'notes-input';
        notesTextarea.id = 'notes-' + task.id;
        notesTextarea.placeholder = 'Add a note or update...';
        section.appendChild(notesTextarea);

        var attachSection = document.createElement('div');
        attachSection.className = 'attachments-section';

        var attachLabel = document.createElement('label');
        attachLabel.textContent = 'Attachments';
        attachSection.appendChild(attachLabel);

        var fileUpload = document.createElement('div');
        fileUpload.className = 'file-upload';
        fileUpload.addEventListener('click', function() {
            var input = this.querySelector('input');
            if (input) input.click();
        });

        var fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.multiple = true;
        fileInput.id = 'file-' + task.id;
        fileInput.addEventListener('change', handleFileSelect);
        fileUpload.appendChild(fileInput);

        var uploadText = document.createElement('div');
        uploadText.className = 'file-upload-text';
        var clickSpan = document.createElement('span');
        clickSpan.textContent = 'Click to upload';
        uploadText.appendChild(clickSpan);
        uploadText.appendChild(document.createTextNode(' or drag and drop'));
        fileUpload.appendChild(uploadText);

        attachSection.appendChild(fileUpload);

        if (task.attachments && task.attachments.length > 0) {
            var fileList = document.createElement('div');
            fileList.className = 'file-list';
            fileList.id = 'files-' + task.id;

            task.attachments.forEach(function(file) {
                var fileItem = document.createElement('div');
                fileItem.className = 'file-item';

                var fileName = document.createElement('span');
                fileName.textContent = file;
                fileItem.appendChild(fileName);

                var removeBtn = document.createElement('span');
                removeBtn.className = 'remove-file';
                removeBtn.textContent = 'Remove';
                removeBtn.addEventListener('click', function(e) {
                    e.stopPropagation();
                    fileItem.remove();
                });
                fileItem.appendChild(removeBtn);

                fileList.appendChild(fileItem);
            });

            attachSection.appendChild(fileList);
        }

        section.appendChild(attachSection);

        var saveBtn = document.createElement('button');
        saveBtn.className = 'save-btn';
        saveBtn.textContent = 'Save Changes';
        saveBtn.dataset.taskId = task.id;
        saveBtn.addEventListener('click', handleSave);
        section.appendChild(saveBtn);

        return section;
    }

    function createInfoSidebar(task) {
        var sidebar = document.createElement('div');
        sidebar.className = 'task-info-sidebar';

        var infoRows = [
            { label: 'Task ID', value: task.id },
            { label: 'Phase', value: task.phase },
            { label: 'Workstream', value: task.workstream },
            { label: 'Due Date', value: formatDate(task.dueDate) },
            { label: 'Status', value: formatStatus(task.status) }
        ];

        infoRows.forEach(function(row) {
            var div = document.createElement('div');
            div.className = 'info-row';

            var label = document.createElement('span');
            label.className = 'info-label';
            label.textContent = row.label;
            div.appendChild(label);

            var value = document.createElement('span');
            value.className = 'info-value';
            value.textContent = row.value;
            div.appendChild(value);

            sidebar.appendChild(div);
        });

        if (task.notes && task.notes.length > 0) {
            var notesHistory = document.createElement('div');
            notesHistory.className = 'notes-history';

            var historyTitle = document.createElement('h5');
            historyTitle.textContent = 'Recent Notes';
            notesHistory.appendChild(historyTitle);

            task.notes.slice(0, 3).forEach(function(note) {
                var entry = document.createElement('div');
                entry.className = 'note-entry';

                var dateDiv = document.createElement('div');
                dateDiv.className = 'note-date';
                dateDiv.textContent = formatDate(note.date);
                entry.appendChild(dateDiv);

                var textDiv = document.createElement('div');
                textDiv.className = 'note-text';
                textDiv.textContent = note.text;
                entry.appendChild(textDiv);

                notesHistory.appendChild(entry);
            });

            sidebar.appendChild(notesHistory);
        }

        return sidebar;
    }

    function toggleTask(taskId) {
        var cards = document.querySelectorAll('.task-card');
        cards.forEach(function(card) {
            if (card.dataset.taskId === taskId) {
                card.classList.toggle('expanded');
                expandedTaskId = card.classList.contains('expanded') ? taskId : null;
            } else {
                card.classList.remove('expanded');
            }
        });
    }

    function handleStatusClick(e) {
        var btn = e.target;
        var taskId = btn.dataset.taskId;
        var newStatus = btn.dataset.status;

        var parent = btn.parentElement;
        parent.querySelectorAll('.status-btn').forEach(function(b) {
            b.classList.remove('active');
        });
        btn.classList.add('active');

        var task = myTasks.find(function(t) { return t.id === taskId; });
        if (task) {
            task.status = newStatus;
            if (newStatus === 'completed') {
                task.progress = 100;
                var slider = document.getElementById('progress-' + taskId);
                var display = document.getElementById('progress-display-' + taskId);
                if (slider) slider.value = '100';
                if (display) display.textContent = '100%';
            }
        }
    }

    function handleProgressChange(e) {
        var slider = e.target;
        var taskId = slider.id.replace('progress-', '');
        var value = slider.value;

        var display = document.getElementById('progress-display-' + taskId);
        if (display) display.textContent = value + '%';

        var task = myTasks.find(function(t) { return t.id === taskId; });
        if (task) task.progress = parseInt(value, 10);
    }

    function handleFileSelect(e) {
        var input = e.target;
        var taskId = input.id.replace('file-', '');
        var maxFileSize = 10 * 1024 * 1024; // 10MB limit
        var allowedTypes = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv', 'png', 'jpg', 'jpeg', 'gif', 'zip'];

        if (!input.files || input.files.length === 0) return;

        var fileList = document.getElementById('files-' + taskId);
        if (!fileList) {
            fileList = document.createElement('div');
            fileList.className = 'file-list';
            fileList.id = 'files-' + taskId;
            // Find the attachments-section parent more robustly
            var attachSection = null;
            if (input.closest) {
                attachSection = input.closest('.attachments-section');
            } else if (input.parentElement && input.parentElement.parentElement) {
                attachSection = input.parentElement.parentElement;
            }
            if (attachSection) attachSection.appendChild(fileList);
        }

        // ES5-compatible iteration over FileList
        var files = [];
        for (var i = 0; i < input.files.length; i++) {
            files.push(input.files[i]);
        }
        files.forEach(function(file) {
            // Validate file size
            if (file.size > maxFileSize) {
                showToast('File "' + file.name + '" exceeds 10MB limit', true);
                return;
            }

            // Validate file type
            var ext = file.name.split('.').pop().toLowerCase();
            if (allowedTypes.indexOf(ext) === -1) {
                showToast('File type ".' + ext + '" not allowed', true);
                return;
            }

            // Check for duplicates
            var existingFiles = fileList.querySelectorAll('.file-item span:first-child');
            for (var i = 0; i < existingFiles.length; i++) {
                if (existingFiles[i].textContent === file.name) {
                    showToast('File "' + file.name + '" already added', true);
                    return;
                }
            }

            var fileItem = document.createElement('div');
            fileItem.className = 'file-item';

            var fileName = document.createElement('span');
            fileName.textContent = file.name;
            fileItem.appendChild(fileName);

            var fileSize = document.createElement('span');
            fileSize.className = 'file-size';
            fileSize.textContent = formatFileSize(file.size);
            fileSize.style.cssText = 'color: #6b7280; font-size: 0.75rem; margin-left: 0.5rem;';
            fileItem.appendChild(fileSize);

            var removeBtn = document.createElement('span');
            removeBtn.className = 'remove-file';
            removeBtn.textContent = 'Remove';
            removeBtn.addEventListener('click', function() { fileItem.remove(); });
            fileItem.appendChild(removeBtn);

            fileList.appendChild(fileItem);
        });
    }

    function handleSave(e) {
        var btn = e.target;
        var taskId = btn.dataset.taskId;

        var notesInput = document.getElementById('notes-' + taskId);
        var newNote = notesInput ? notesInput.value.trim() : '';

        var task = myTasks.find(function(t) { return t.id === taskId; });
        if (task && newNote) {
            task.notes.unshift({
                date: new Date().toISOString().split('T')[0],
                text: newNote
            });
            notesInput.value = '';
        }

        showToast('Task updated successfully!');
        updateStats();

        var card = document.querySelector('[data-task-id="' + taskId + '"]');
        if (card && task) {
            var indicator = card.querySelector('.task-status-indicator');
            if (indicator) {
                indicator.className = 'task-status-indicator ' + task.status;
            }
        }
    }

    function setupFilters() {
        var tabs = document.querySelectorAll('.filter-tab');
        tabs.forEach(function(tab) {
            tab.addEventListener('click', function() {
                tabs.forEach(function(t) { t.classList.remove('active'); });
                tab.classList.add('active');
                currentFilter = tab.dataset.filter;
                renderTasks();
            });
        });
    }

    function setupSearch() {
        var input = document.getElementById('searchInput');
        if (input) {
            input.addEventListener('input', debounce(function(e) {
                searchTerm = e.target.value;
                renderTasks();
            }, 300));
        }
    }

    function updateFooterDate() {
        var dateEl = document.getElementById('currentDate');
        if (dateEl) {
            dateEl.textContent = new Date().toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'long',
                day: 'numeric'
            });
        }
    }

    function formatDate(dateStr) {
        if (!dateStr) return 'N/A';
        var date = new Date(dateStr);
        // Check for Invalid Date
        if (isNaN(date.getTime())) return 'N/A';
        return date.toLocaleDateString('en-US', {
            month: 'short',
            day: 'numeric',
            year: 'numeric'
        });
    }

    function formatStatus(status) {
        var statusMap = {
            'not-started': 'Not Started',
            'in-progress': 'In Progress',
            'on-track': 'On Track',
            'at-risk': 'At Risk',
            'blocked': 'Blocked',
            'completed': 'Completed'
        };
        return statusMap[status] || status;
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        var k = 1024;
        var sizes = ['B', 'KB', 'MB', 'GB'];
        var i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    function showToast(message, isError) {
        var toast = document.getElementById('toast');
        if (!toast) return;

        toast.textContent = message;
        toast.className = 'toast' + (isError ? ' error' : '');
        toast.classList.add('show');

        setTimeout(function() {
            toast.classList.remove('show');
        }, 3000);
    }

    function debounce(func, wait) {
        var timeout;
        return function() {
            var context = this;
            var args = arguments;
            var later = function() {
                clearTimeout(timeout);
                func.apply(context, args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
