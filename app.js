class XMLDiffer {
    constructor() {
        this.file1Content = null;
        this.file2Content = null;
        this.file1Name = '';
        this.file2Name = '';
        this.currentDiff = null;
        this.viewMode = 'unified'; // 'unified' or 'split'
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const file1Input = document.getElementById('file1');
        const file2Input = document.getElementById('file2');
        const file1InputDiv = document.getElementById('file1-input');
        const file2InputDiv = document.getElementById('file2-input');
        const compareBtn = document.getElementById('compare-btn');
        const unifiedBtn = document.getElementById('unified-btn');
        const splitBtn = document.getElementById('split-btn');

        // File input handling
        file1InputDiv.addEventListener('click', () => file1Input.click());
        file2InputDiv.addEventListener('click', () => file2Input.click());

        file1Input.addEventListener('change', (e) => this.handleFileSelect(e, 1));
        file2Input.addEventListener('change', (e) => this.handleFileSelect(e, 2));

        // Drag and drop
        this.setupDragAndDrop(file1InputDiv, file1Input, 1);
        this.setupDragAndDrop(file2InputDiv, file2Input, 2);

        compareBtn.addEventListener('click', () => this.compareFiles());

        // View mode toggle
        unifiedBtn.addEventListener('click', () => this.switchView('unified'));
        splitBtn.addEventListener('click', () => this.switchView('split'));
    }

    setupDragAndDrop(element, input, fileNum) {
        element.addEventListener('dragover', (e) => {
            e.preventDefault();
            element.style.borderColor = '#0366d6';
            element.style.background = '#f1f8ff';
        });

        element.addEventListener('dragleave', () => {
            element.style.borderColor = '#d1d5da';
            element.style.background = '#fafbfc';
        });

        element.addEventListener('drop', (e) => {
            e.preventDefault();
            element.style.borderColor = '#d1d5da';
            element.style.background = '#fafbfc';

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                input.files = files;
                this.handleFileSelect({ target: input }, fileNum);
            }
        });
    }

    async handleFileSelect(event, fileNum) {
        const file = event.target.files[0];
        if (!file) return;

        if (!file.name.toLowerCase().endsWith('.xml')) {
            alert('Please select an XML file');
            return;
        }

        try {
            const content = await this.readFile(file);
            const label = document.getElementById(`file${fileNum}-label`);
            const inputDiv = document.getElementById(`file${fileNum}-input`);

            if (fileNum === 1) {
                this.file1Content = content;
                this.file1Name = file.name;
            } else {
                this.file2Content = content;
                this.file2Name = file.name;
            }

            label.textContent = `✓ ${file.name}`;
            inputDiv.classList.add('has-file');

            this.updateCompareButton();
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading file. Please try again.');
        }
    }

    readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsText(file);
        });
    }

    updateCompareButton() {
        const compareBtn = document.getElementById('compare-btn');
        compareBtn.disabled = !(this.file1Content && this.file2Content);
    }

    async compareFiles() {
        if (!this.file1Content || !this.file2Content) return;

        const diffContainer = document.getElementById('diff-container');
        const unifiedContainer = document.getElementById('unified-container');
        const leftPane = document.getElementById('left-pane');
        const rightPane = document.getElementById('right-pane');

        // Show loading
        unifiedContainer.innerHTML = '<div class="loading">Processing...</div>';
        leftPane.innerHTML = '<div class="loading">Processing...</div>';
        rightPane.innerHTML = '<div class="loading">Processing...</div>';
        diffContainer.style.display = 'block';

        try {
            // Format XML for better comparison
            const formatted1 = this.formatXML(this.file1Content);
            const formatted2 = this.formatXML(this.file2Content);

            // Calculate diff
            this.currentDiff = this.calculateDiff(formatted1, formatted2);

            // Update file info
            this.updateFileInfo();

            // Render current view
            this.renderCurrentView();
            this.updateStats(this.currentDiff);

        } catch (error) {
            console.error('Error comparing files:', error);
            const errorMsg = '<div class="error">Error processing files. Please check that both files are valid XML.</div>';
            unifiedContainer.innerHTML = errorMsg;
            leftPane.innerHTML = errorMsg;
            rightPane.innerHTML = errorMsg;
        }
    }

    switchView(mode) {
        this.viewMode = mode;

        const unifiedBtn = document.getElementById('unified-btn');
        const splitBtn = document.getElementById('split-btn');
        const diffContent = document.getElementById('diff-content');

        // Update button states
        unifiedBtn.classList.toggle('active', mode === 'unified');
        splitBtn.classList.toggle('active', mode === 'split');

        // Update view class
        diffContent.className = `diff-content ${mode === 'unified' ? 'unified-view' : 'split-view'}`;

        // Show/hide appropriate panes
        const unifiedPane = document.getElementById('unified-pane');
        const leftPaneContainer = document.getElementById('left-pane-container');
        const rightPaneContainer = document.getElementById('right-pane-container');

        if (mode === 'unified') {
            unifiedPane.style.display = 'block';
            leftPaneContainer.style.display = 'none';
            rightPaneContainer.style.display = 'none';
        } else {
            unifiedPane.style.display = 'none';
            leftPaneContainer.style.display = 'block';
            rightPaneContainer.style.display = 'block';
        }

        // Re-render if we have diff data
        if (this.currentDiff) {
            this.renderCurrentView();
        }
    }

    updateFileInfo() {
        const leftInfo = document.getElementById('left-info');
        const rightInfo = document.getElementById('right-info');
        const unifiedInfo = document.getElementById('unified-info');

        leftInfo.textContent = this.file1Name;
        rightInfo.textContent = this.file2Name;
        unifiedInfo.textContent = `${this.file1Name} ↔ ${this.file2Name}`;
    }

    renderCurrentView() {
        if (!this.currentDiff) return;

        if (this.viewMode === 'unified') {
            this.renderUnifiedView(this.currentDiff);
        } else {
            this.renderSplitView(this.currentDiff);
        }
    }

    formatXML(xmlString) {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlString, 'text/xml');

            // Check for parsing errors
            const parserError = xmlDoc.getElementsByTagName('parsererror');
            if (parserError.length > 0) {
                throw new Error('Invalid XML');
            }

            return this.xmlToString(xmlDoc, 0).split('\n').filter(line => line.trim());
        } catch (error) {
            // If XML parsing fails, return original split by lines
            return xmlString.split('\n');
        }
    }

    xmlToString(node, indent = 0) {
        const indentStr = '  '.repeat(indent);
        let result = '';

        if (node.nodeType === Node.ELEMENT_NODE) {
            result += indentStr + '<' + node.nodeName;

            // Add attributes
            for (let attr of node.attributes || []) {
                result += ` ${attr.name}="${attr.value}"`;
            }

            if (node.childNodes.length === 0) {
                result += '/>\n';
            } else {
                result += '>\n';

                for (let child of node.childNodes) {
                    if (child.nodeType === Node.TEXT_NODE) {
                        const text = child.textContent.trim();
                        if (text) {
                            result += '  '.repeat(indent + 1) + text + '\n';
                        }
                    } else if (child.nodeType === Node.ELEMENT_NODE) {
                        result += this.xmlToString(child, indent + 1);
                    } else if (child.nodeType === Node.COMMENT_NODE) {
                        result += '  '.repeat(indent + 1) + '<!-- ' + child.textContent + ' -->\n';
                    }
                }

                result += indentStr + '</' + node.nodeName + '>\n';
            }
        } else if (node.nodeType === Node.DOCUMENT_NODE) {
            for (let child of node.childNodes) {
                result += this.xmlToString(child, indent);
            }
        }

        return result;
    }

    calculateDiff(lines1, lines2) {
        const diff = [];
        const lcs = this.getLCS(lines1, lines2);

        let i = 0, j = 0, lcsIndex = 0;

        while (i < lines1.length || j < lines2.length) {
            if (lcsIndex < lcs.length && i < lines1.length && j < lines2.length &&
                lines1[i] === lcs[lcsIndex] && lines2[j] === lcs[lcsIndex]) {
                // Unchanged line
                diff.push({
                    type: 'unchanged',
                    left: { lineNum: i + 1, content: lines1[i] },
                    right: { lineNum: j + 1, content: lines2[j] }
                });
                i++;
                j++;
                lcsIndex++;
            } else if (i < lines1.length && (lcsIndex >= lcs.length || lines1[i] !== lcs[lcsIndex])) {
                // Removed line
                diff.push({
                    type: 'removed',
                    left: { lineNum: i + 1, content: lines1[i] },
                    right: { lineNum: null, content: '' }
                });
                i++;
            } else if (j < lines2.length && (lcsIndex >= lcs.length || lines2[j] !== lcs[lcsIndex])) {
                // Added line
                diff.push({
                    type: 'added',
                    left: { lineNum: null, content: '' },
                    right: { lineNum: j + 1, content: lines2[j] }
                });
                j++;
            }
        }

        return this.identifyModifiedLines(diff);
    }

    identifyModifiedLines(diff) {
        const result = [];
        let i = 0;

        while (i < diff.length) {
            if (i < diff.length - 1 &&
                diff[i].type === 'removed' &&
                diff[i + 1].type === 'added') {

                // Check if lines are similar (potential modification)
                const similarity = this.calculateSimilarity(diff[i].left.content, diff[i + 1].right.content);

                if (similarity > 0.5) {
                    // Treat as modified line
                    result.push({
                        type: 'modified',
                        left: diff[i].left,
                        right: diff[i + 1].right
                    });
                    i += 2;
                } else {
                    result.push(diff[i]);
                    i++;
                }
            } else {
                result.push(diff[i]);
                i++;
            }
        }

        return result;
    }

    calculateSimilarity(str1, str2) {
        const longer = str1.length > str2.length ? str1 : str2;
        const shorter = str1.length > str2.length ? str2 : str1;

        if (longer.length === 0) return 1.0;

        const editDistance = this.levenshteinDistance(longer, shorter);
        return (longer.length - editDistance) / longer.length;
    }

    levenshteinDistance(str1, str2) {
        const matrix = [];

        for (let i = 0; i <= str2.length; i++) {
            matrix[i] = [i];
        }

        for (let j = 0; j <= str1.length; j++) {
            matrix[0][j] = j;
        }

        for (let i = 1; i <= str2.length; i++) {
            for (let j = 1; j <= str1.length; j++) {
                if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        matrix[i][j - 1] + 1,
                        matrix[i - 1][j] + 1
                    );
                }
            }
        }

        return matrix[str2.length][str1.length];
    }

    getLCS(arr1, arr2) {
        const m = arr1.length;
        const n = arr2.length;
        const dp = Array(m + 1).fill().map(() => Array(n + 1).fill(0));

        for (let i = 1; i <= m; i++) {
            for (let j = 1; j <= n; j++) {
                if (arr1[i - 1] === arr2[j - 1]) {
                    dp[i][j] = dp[i - 1][j - 1] + 1;
                } else {
                    dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
                }
            }
        }

        // Reconstruct LCS
        const lcs = [];
        let i = m, j = n;

        while (i > 0 && j > 0) {
            if (arr1[i - 1] === arr2[j - 1]) {
                lcs.unshift(arr1[i - 1]);
                i--;
                j--;
            } else if (dp[i - 1][j] > dp[i][j - 1]) {
                i--;
            } else {
                j--;
            }
        }

        return lcs;
    }

    renderUnifiedView(diff) {
        const container = document.getElementById('unified-container');
        let html = '';

        diff.forEach(item => {
            const className = `unified-line line-${item.type}`;

            if (item.type === 'unchanged') {
                html += this.renderUnifiedLine(item.left.lineNum, item.right.lineNum, item.left.content, className);
            } else if (item.type === 'added') {
                html += this.renderUnifiedLine('', item.right.lineNum, item.right.content, className);
            } else if (item.type === 'removed') {
                html += this.renderUnifiedLine(item.left.lineNum, '', item.left.content, className);
            } else if (item.type === 'modified') {
                // Show both old and new versions
                html += this.renderUnifiedLine(item.left.lineNum, '', item.left.content, 'unified-line line-removed');
                html += this.renderUnifiedLine('', item.right.lineNum, item.right.content, 'unified-line line-added');
            }
        });

        container.innerHTML = html;
    }

    renderUnifiedLine(leftNum, rightNum, content, className) {
        const highlightedContent = this.highlightXML(content || '');
        const hasLeftNum = leftNum !== '';
        const hasRightNum = rightNum !== '';

        // Add visual context about which file the line comes from
        let lineContext = '';
        if (hasLeftNum && !hasRightNum) {
            lineContext = '<span class="file-indicator file-indicator-a" style="margin-right: 8px;">A</span>';
        } else if (!hasLeftNum && hasRightNum) {
            lineContext = '<span class="file-indicator file-indicator-b" style="margin-right: 8px;">B</span>';
        }

        return `<div class="${className}">
                    <div class="line-numbers">
                        <div class="line-num-left">${leftNum}</div>
                        <div class="line-num-right">${rightNum}</div>
                    </div>
                    <div class="line-content line-context">${lineContext}${highlightedContent}</div>
                </div>`;
    }

    renderSplitView(diff) {
        const leftPane = document.getElementById('left-pane');
        const rightPane = document.getElementById('right-pane');

        let leftHTML = '';
        let rightHTML = '';

        diff.forEach(item => {
            const leftClass = this.getLineClass(item.type, 'left');
            const rightClass = this.getLineClass(item.type, 'right');

            leftHTML += this.renderLine(item.left, leftClass);
            rightHTML += this.renderLine(item.right, rightClass);
        });

        leftPane.innerHTML = leftHTML;
        rightPane.innerHTML = rightHTML;
    }

    renderDiff(diff) {
        this.renderSplitView(diff);
    }

    getLineClass(type, side) {
        switch (type) {
            case 'added': return side === 'left' ? 'line-unchanged' : 'line-added';
            case 'removed': return side === 'left' ? 'line-removed' : 'line-unchanged';
            case 'modified': return 'line-modified';
            default: return 'line-unchanged';
        }
    }

    renderLine(line, className) {
        const lineNum = line.lineNum || '';
        const content = this.highlightXML(line.content || '');

        return `<div class="line ${className}">
                    <div class="line-number">${lineNum}</div>
                    <div class="line-content">${content}</div>
                </div>`;
    }

    highlightXML(content) {
        if (!content.trim()) return '&nbsp;';

        return content
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/(&lt;\/?[^&\s]+)([^&]*?)(&gt;)/g, '<span class="xml-tag">$1$2$3</span>')
            .replace(/(\w+)=("[^"]*")/g, '<span class="xml-attr-name">$1</span>=<span class="xml-attr-value">$2</span>')
            .replace(/(&lt;!--.*?--&gt;)/g, '<span class="xml-comment">$1</span>');
    }

    updateStats(diff) {
        const stats = diff.reduce((acc, item) => {
            acc[item.type] = (acc[item.type] || 0) + 1;
            return acc;
        }, {});

        const statsContainer = document.getElementById('diff-stats');
        let statsHTML = '';

        if (stats.added) {
            statsHTML += `<div class="stat-item stat-added"><span class="file-indicator file-indicator-b">B</span>+${stats.added} added</div>`;
        }
        if (stats.removed) {
            statsHTML += `<div class="stat-item stat-removed"><span class="file-indicator file-indicator-a">A</span>-${stats.removed} removed</div>`;
        }
        if (stats.modified) {
            statsHTML += `<div class="stat-item stat-modified">${stats.modified} modified</div>`;
        }
        if (stats.unchanged) {
            statsHTML += `<div class="stat-item">${stats.unchanged} unchanged</div>`;
        }

        statsContainer.innerHTML = statsHTML;
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    new XMLDiffer();
});
