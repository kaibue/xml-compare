let xml1 = '', xml2 = '';

document.getElementById('file1').addEventListener('change', function () {
    readFile(this, 1);
});
document.getElementById('file2').addEventListener('change', function () {
    readFile(this, 2);
});

function readFile(input, index) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
        if (index === 1) xml1 = reader.result;
        else xml2 = reader.result;
        if (xml1 && xml2) showDiff(xml1, xml2);
    };
    reader.readAsText(file);
}

function escapeHTML(str) {
    return str.replace(/[&<>"']/g, (m) => {
        switch (m) {
            case '&': return '&amp;';
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '"': return '&quot;';
            case "'": return '&#39;';
        }
    });
}

function highlightWordDiffs(line1, line2) {
    const tokenize = (line) =>
        line.match(/<\/?[\w:-]+|\/?>|[\w:-]+="[^"]*"|[\w:-]+|".*?"|\s+|[^\w\s]/g) || [];

    const tokens1 = tokenize(line1);
    const tokens2 = tokenize(line2);

    let html = '';
    let inHighlight = false;

    const isTag = (t) => /^<\/?[\w:-]+$/.test(t);
    const isSelfClose = (t) => /^\/?>$/.test(t);
    const isAttribute = (t) => /^[\w:-]+="[^"]*"$/.test(t);

    const extractTagName = (t) => t.replace(/^<\/?/, '');

    const splitAttribute = (attr) => {
        const match = attr.match(/^([\w:-]+)=(")([^"]*)(")$/);
        if (!match) return null;
        return { name: match[1], quote1: match[2], value: match[3], quote2: match[4] };
    };

    function diffInnerText(text1, text2) {
        let result = '';
        let inDiff = false;
        const maxLen = Math.max(text1.length, text2.length);
        for (let i = 0; i < maxLen; i++) {
            const c1 = text1[i] || '';
            const c2 = text2[i] || '';

            if (c1 === c2) {
                if (inDiff) {
                    result += '</span>';
                    inDiff = false;
                }
                result += escapeHTML(c1);
            } else {
                if (!inDiff) {
                    result += '<span class="diff-change-word" title="Word changed">';
                    inDiff = true;
                }
                result += escapeHTML(c1);
            }
        }
        if (inDiff) result += '</span>';
        return result;
    }

    for (let i = 0; i < tokens1.length; i++) {
        const t1 = tokens1[i];
        const t2 = tokens2[i] || '';

        if (isTag(t1) && isTag(t2)) {
            if (t1 === t2) {
                if (inHighlight) {
                    html += '</span>';
                    inHighlight = false;
                }
                html += escapeHTML(t1);
            } else {
                if (!inHighlight) {
                    html += '<span class="diff-change-word" title="Tag changed">';
                    inHighlight = true;
                }
                html += escapeHTML(t1);
            }
        } else if (isAttribute(t1) && isAttribute(t2)) {
            const attr1 = splitAttribute(t1);
            const attr2 = splitAttribute(t2);

            if (attr1 && attr2 && attr1.name === attr2.name) {
                if (inHighlight) {
                    html += '</span>';
                    inHighlight = false;
                }
                html += escapeHTML(attr1.name) + '=' + attr1.quote1;
                html += diffInnerText(attr1.value, attr2.value);
                html += attr1.quote2;
            } else {
                if (!inHighlight) {
                    html += '<span class="diff-change-word" title="Attribute changed">';
                    inHighlight = true;
                }
                html += escapeHTML(t1);
            }
        } else if (!isTag(t1) && !isSelfClose(t1) && !isAttribute(t1)) {
            if (t1 === t2) {
                if (inHighlight) {
                    html += '</span>';
                    inHighlight = false;
                }
                html += escapeHTML(t1);
            } else {
                if (!inHighlight) {
                    html += '<span class="diff-change-word" title="Text changed">';
                    inHighlight = true;
                }
                html += escapeHTML(t1);
            }
        } else {
            if (inHighlight) {
                html += '</span>';
                inHighlight = false;
            }
            html += escapeHTML(t1);
        }
    }

    if (inHighlight) html += '</span>';
    return html;
}

function alignLinesWithPadding(origA, origB) {
    const output = [];
    let i = 0, j = 0;

    while (i < origA.length || j < origB.length) {
        const lineA = origA[i];
        const lineB = origB[j];
        const normLineA = (lineA || '').trim();
        const normLineB = (lineB || '').trim();

        const currentLineNumA = i + 1;
        const currentLineNumB = j + 1;

        // Skip identical blank lines
        if (normLineA === '' && normLineB === '' && i < origA.length && j < origB.length) {
            i++;
            j++;
            continue; // Skip this pair and move to the next lines
        }

        if (normLineA === normLineB) {
            output.push({ lineA: lineA, lineB: lineB, originalLineA: currentLineNumA, originalLineB: currentLineNumB });
            i++;
            j++;
        } else if (!normLineA) {
            output.push({ lineA: '', lineB: lineB, originalLineA: '', originalLineB: currentLineNumB });
            j++;
        } else if (!normLineB) {
            output.push({ lineA: lineA, lineB: '', originalLineA: currentLineNumA, originalLineB: '' });
            i++;
        } else {
            let foundMatchInB = -1;
            let foundMatchInA = -1;

            // Look ahead in B for a match with lineA
            for (let k = j + 1; k < origB.length; k++) {
                if (normLineA === origB[k].trim()) {
                    foundMatchInB = k;
                    break;
                }
            }

            // Look ahead in A for a match with lineB
            for (let k = i + 1; k < origA.length; k++) {
                if (normLineB === origA[k].trim()) {
                    foundMatchInA = k;
                    break;
                }
            }

            if (foundMatchInB !== -1 && (foundMatchInA === -1 || foundMatchInB - j <= foundMatchInA - i)) {
                // Prefer insertions in B if match is found sooner or no match in A
                while (j < foundMatchInB) {
                    const nextLineB = origB[j];
                    output.push({ lineA: '', lineB: nextLineB, originalLineA: '', originalLineB: j + 1 });
                    j++;
                }
            } else if (foundMatchInA !== -1) {
                // Prefer insertions in A
                while (i < foundMatchInA) {
                    const nextLineA = origA[i];
                    output.push({ lineA: nextLineA, lineB: '', originalLineA: i + 1, originalLineB: '' });
                    i++;
                }
            } else {
                // No clear match, treat as a change
                output.push({ lineA: lineA, lineB: lineB, originalLineA: currentLineNumA, originalLineB: currentLineNumB });
                i++;
                j++;
            }
        }
    }

    return output;
}

function showDiff(text1, text2) {
    const lines1 = text1.split(/\r?\n/);
    const lines2 = text2.split(/\r?\n/);

    const diffs = alignLinesWithPadding(lines1, lines2);

    let html1 = '', html2 = '';

    diffs.forEach(diffItem => {
        const { lineA, lineB, originalLineA, originalLineB } = diffItem;

        // Skip rendering if both lines are blank after processing
        if (lineA.trim() === '' && lineB.trim() === '') {
            return;
        }

        const maxLen = Math.max(lineA.length, lineB.length);
        const paddedL1 = lineA.padEnd(maxLen, ' ');
        const paddedL2 = lineB.padEnd(maxLen, ' ');

        let class1 = '', class2 = '';
        let content1 = escapeHTML(paddedL1);
        let content2 = escapeHTML(paddedL2);
        let title1 = '';
        let title2 = '';

        if (lineA.trim() === lineB.trim()) {
            class1 = class2 = '';
        } else if (!lineA.trim()) {
            class1 = 'diff-remove';
            class2 = 'diff-add';
            title1 = 'Line missing in File 1';
            title2 = 'Line added in File 2';
        } else if (!lineB.trim()) {
            class1 = 'diff-add';
            class2 = 'diff-remove';
            title1 = 'Line added in File 1';
            title2 = 'Line missing in File 2';
        } else {
            class1 = class2 = 'diff-change';
            title1 = title2 = 'Line changed';
            content1 = highlightWordDiffs(lineA, lineB);
            content2 = highlightWordDiffs(lineB, lineA);
        }

        const lineNumber1 = originalLineA !== '' ? originalLineA : '';
        const lineNumber2 = originalLineB !== '' ? originalLineB : '';

        html1 += `<div class="diff-line-wrapper"><span class="line-number">${lineNumber1}</span><div class="${class1}" title="${title1}">${content1}</div></div>`;
        html2 += `<div class="diff-line-wrapper"><span class="line-number">${lineNumber2}</span><div class="${class2}" title="${title2}">${content2}</div></div>`;
    });

    document.getElementById('output1').innerHTML = html1;
    document.getElementById('output2').innerHTML = html2;
}

// Sync scrolling
const output1 = document.getElementById('output1');
const output2 = document.getElementById('output2');
function syncScroll(source, target) {
    target.scrollTop = source.scrollTop;
    target.scrollLeft = source.scrollLeft;
}
output1.addEventListener('scroll', () => syncScroll(output1, output2));
output2.addEventListener('scroll', () => syncScroll(output2, output1));
