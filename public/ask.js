async function callGemini(prompt, tableData) {
    const API_KEY = "AIzaSyD_F5y16ZA1Zd1m33anlfZVj-ortAs_ifQ"; // your key
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${API_KEY}`;

    const requestBody = {
        contents: [{
            parts: [{
                text: `You are an Excel AI assistant. 
                Prompt: ${prompt} 
                Data: ${JSON.stringify(tableData)}`
            }]
        }]
    };

    const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
    });

    if (!res.ok) {
        throw new Error(`Gemini API error: ${await res.text()}`);
    }

    const data = await res.json();
    return data.candidates[0].content.parts[0].text;
}

// Example: Get selected table in Excel
async function runAgent() {
    await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("values");
        await context.sync();

        const prompt = document.getElementById("promptInput").value;
        const tableData = range.values;

        const output = await callGemini(prompt, tableData);
        document.getElementById("output").value = output;
    });
}

document.getElementById("runBtn").addEventListener("click", runAgent);
