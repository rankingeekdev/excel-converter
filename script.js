function convertToExcel() {
    // Get the question paper name and input text
    const paperName = document.getElementById("paperName").value.trim();
    const input = document.getElementById("questionInput").value;

    // Parse questions from input
    const questions = parseQuestions(input, paperName);

    // Create a worksheet and workbook
    const ws = XLSX.utils.json_to_sheet(questions);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Questions");

    // Generate Excel file and trigger download
    XLSX.writeFile(wb, `${paperName || "questions"}.xlsx`);
}

function parseQuestions(input, paperName) {
    const lines = input.split('\n');
    const questions = [];
    let currentQuestion = null;
    let options = [];

    lines.forEach(line => {
        // Trim the line to remove extra spaces
        line = line.trim();

        // Check for a new question (lines starting with #)
        if (line.startsWith('#')) {
            // If there's an existing question, push it to the questions array
            if (currentQuestion) {
                // Ensure we have 4 options, filling empty slots if necessary
                while (options.length < 4) {
                    options.push('');
                }
                currentQuestion.options = options;
                questions.push(currentQuestion);
            }

            // Start a new question and reset the options array
            currentQuestion = { question: line, options: [] };
            options = [];
        } else if (line && options.length < 4) {
            // Add up to 4 options (avoid capturing extra lines)
            options.push(line);
        }
    });

    // Push the last question, if exists
    if (currentQuestion) {
        while (options.length < 4) {
            options.push('');
        }
        currentQuestion.options = options;
        questions.push(currentQuestion);
    }

    // Map questions and options into a format for Excel export
    return questions.map(q => ({
        "Question Paper Name": paperName || "Untitled Paper",
        Question: q.question,
        OptionA: q.options[0] || '',
        OptionB: q.options[1] || '',
        OptionC: q.options[2] || '',
        OptionD: q.options[3] || ''
    }));
}
