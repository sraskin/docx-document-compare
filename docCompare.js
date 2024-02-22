const fs = require('fs'); // A library to read and write files
const diff = require('diff'); // A library to compare text files
const mammoth = require('mammoth'); // A .docx file reader

/**
 * Reads a .docx file and extracts its text content.
 * @param {string} file - The path to the .docx file.
 * @returns {Promise<string>} The text content of the .docx file.
 */
async function readDocx(file) {
    const buffer = fs.readFileSync(file);
    const { value: text } = await mammoth.extractRawText({ buffer: buffer });
    return text;
}

/**
 * Compares the text content of two .docx files and returns the differences.
 * @param {string} file1 - The path to the first .docx file.
 * @param {string} file2 - The path to the second .docx file.
 * @param {string} mistakeType - The type of mistake to be logged.
 * @returns {Promise<Array>} An array of mistakes found between the two files.
 */
async function compareFiles(file1, file2, mistakeType) {
    const text1 = await readDocx(file1);
    const text2 = await readDocx(file2);

    const lines1 = text1.split('\n');
    const lines2 = text2.split('\n');

    let mistakes = [];

    const diffs = diff.diffLines(text1, text2);
    let lineNumber = 1;

    diffs.forEach((part) => {
        // If a part is added or removed, log it as a mistake
        if (part.added || part.removed) {
            const changes = part.value.split('\n');
            changes.pop(); // remove last empty line
            changes.forEach((change) => {
                mistakes.push({
                    file1: file1,
                    file2: file2,
                    lineNumber: lineNumber,
                    characterNumber: lines1[lineNumber - 1].indexOf(change) + 1,
                    text: change,
                    mistakeType: mistakeType
                });
                lineNumber++;
            });
        } else {
            lineNumber += part.count;
        }
    });

    return mistakes;
}

/**
 * Main function to compare multiple pairs of .docx files and log the differences.
 */
async function main() {
    const mistakes1 = await compareFiles('./docs/Client Doc.docx', './docs/Conditional Coupon Doc.docx', 'During Implementation');
    const mistakes2 = await compareFiles('./docs/Guaranteed Coupon Doc.docx', './docs/Latest Guaranteed Coupon Doc.docx', 'After Implementation');

    const result = {
        'Client Doc vs Conditional Coupon Doc': mistakes1,
        'Guaranteed Coupon Doc vs Latest Guaranteed Coupon Doc': mistakes2
    };

    console.log(JSON.stringify(result, null, 2));
}

main();