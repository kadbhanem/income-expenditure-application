
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel to Table & Word</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/docx@8.0.2/build/index.js"></script>


    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            padding: 20px;
            text-align: center;
            background-color: #f4f4f4;
        }
        input, button {
            margin-bottom: 20px;
            padding: 10px;
            font-size: 16px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }
        th, td {
            border: 1px solid #ddd;
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #007bff;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <h2>Read Excel, Display in Table & Convert to Word <button id="downloadBtn">Download Sample Excel sheet</button></h2>
    
    <input type="file" id="fileInput" accept=".xlsx">
    <button id="convertButton" disabled>Convert to Word</button>
    <table id="dataTable">
        <thead>
            <tr>
                <th>Case Name</th>
                <th>Party Name</th>
                <th>Date of Decision</th>
                <th>Judge</th>
                <th>Clerk Name</th>
                <th>Clerk Designation</th>
                <th>Checked By</th>
            </tr>
        </thead>
        <tbody>
            <!-- Data will be inserted here -->
        </tbody>
    </table>

    <script>
        let jsonData = [];

        document.getElementById('fileInput').addEventListener('change', function (event) {
            const file = event.target.files[0];
            if (!file) {
                alert("Please select a file!");
                return;
            }
            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                displayData(jsonData);
                document.getElementById("convertButton").disabled = false;
            };
            reader.readAsArrayBuffer(file);
        });

        function displayData(data) {
            const tableBody = document.querySelector("#dataTable tbody");
            tableBody.innerHTML = "";
            for (let i = 1; i < data.length; i++) { // Skip header row
                let row = document.createElement("tr");
                data[i].forEach(cell => {
                    let cellElement = document.createElement("td");
                    cellElement.textContent = cell;
                    row.appendChild(cellElement);
                });
                tableBody.appendChild(row);
            }
        }

        document.getElementById("convertButton").addEventListener("click", async function () {
            if (jsonData.length < 2) {
                alert("No valid data to convert!");
                return;
            }
            await generateWordDocument(jsonData.slice(1));
        });

        async function generateWordDocument(data) {

                const { Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak } = window.docx;

                if (!data || data.length === 0) {
                    alert("No valid data to generate Word document.");
                    return;
                }

                let sections = data.map((row, index) => {
                    if (!row || row.length < 3) return { properties: {}, children: [new Paragraph("")] }; // Prevent errors

                    let caseName = row[0] || "Case Name";
                    let partyDetails = row[1] ? row[1].split(" Vs ") : ["Petitioner", "Respondent"];
                    let decisionDate = row[2] || "01-01-2025";
                    let judgename=row[3] || "";
                    let clerk_name=row[4] || "";
                    let clerk_desig=row[5] || "";
                    let checked_by=row[6] || "";

                    return {
                        properties: {},
                        children: [
                            new Paragraph({
                                children: [new TextRun({ text: "(B+C+D)     ", bold: true, size: 35 })],
                                alignment: AlignmentType.RIGHT,
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: "फाईल :", bold: true, size: 52 })],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 200 },
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: judgename, bold: true, size: 36 })],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 200 },
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: "", bold: true, size: 30 })],
                                alignment: AlignmentType.CENTER,
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: caseName, bold: true, size: 40 })],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 300 },
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: "", bold: true, size: 30 })],
                                alignment: AlignmentType.CENTER,
                            }),
                           new docx.Paragraph({
					    spacing: { after: 200 },
					    // 1. Define two tab stops: CENTER at the midpoint, RIGHT at the margin
					    tabStops: [
						{
						    type: docx.TabStopType.CENTER,      // Center Alignment
						    position: 4500,                     // Middle of a standard A4 page (adjust if needed)
						},
						{
						    type: docx.TabStopType.RIGHT,       // Right Alignment
						    position: 9000,                     // Near the right margin (adjust if needed)
						},
					    ],
					    children: [
						// Tab 1: Moves cursor to the CENTER tab stop
						new docx.TextRun({ text: "\t" }), 
						
						// Text 1: The party name, centered at 4500
						new docx.TextRun({
						    text: partyDetails[0],
						    bold: true,
						    size: 30,
						}),
						
						// Tab 2: Moves cursor to the RIGHT tab stop
						new docx.TextRun({ text: "\t" }),
						
						// Text 2: The descriptive text, aligned to the right margin at 9000
						new docx.TextRun({
						    text: "(अर्जदार/फिर्यादी)",
						    bold: true,
						    size: 30,
						}),
					    ],
					    // Remove alignment: AlignmentType.CENTER from the Paragraph itself.
					}),
                            new Paragraph({
                                children: [new TextRun({ text: "Vs", bold: true, size: 32 })],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 200 },
                            }),
                            new docx.Paragraph({
				    spacing: { after: 600 },
				    // 1. Define two tab stops: CENTER at the midpoint, RIGHT at the margin
				    tabStops: [
					{
					    type: docx.TabStopType.CENTER,      // Center Alignment
					    position: 4500,                     // Middle of a standard A4 page (adjust if needed)
					},
					{
					    type: docx.TabStopType.RIGHT,       // Right Alignment
					    position: 9000,                     // Near the right margin (adjust if needed)
					},
				    ],
				    children: [
					// Tab 1: Moves cursor to the CENTER tab stop
					new docx.TextRun({ text: "\t" }), 
					
					// Text 1: The party name, centered at 4500
					new docx.TextRun({
					    text: partyDetails[1],
					    bold: true,
					    size: 30,
					}),
					
					// Tab 2: Moves cursor to the RIGHT tab stop
					new docx.TextRun({ text: "\t" }),
					
					// Text 2: The descriptive text, aligned to the right margin at 9000
					new docx.TextRun({
					    text: "(सा. वाले/आरोपी)",
					    bold: true,
					    size: 30,
					}),
				    ],
				    // Remove alignment: AlignmentType.CENTER from the Paragraph itself.
				}),
                            new Paragraph({
                                spacing: { after: 300 },
                                children: [new TextRun({ text: "निकल तारीख : " + decisionDate, bold: true, size: 30 })],
                            }),
                            new Paragraph({ 
                                spacing: { after: 1500 },
                                children: [new TextRun({ text: "निशाणी :", bold: true, size: 30 })] }),
                            new Paragraph({ 
                                spacing: { after: 1000 },
                                children: [
                                    new TextRun({ text: "पान नंबर :", bold: true, size: 30 }),
                                    new TextRun({ text: "    " }),  // This adds a space after "पान नंबर १ ते"
                                    new TextRun({ text: "     १ ते",bold: true, size: 30 })  // This adds a space after "पान नंबर १ ते"
                                ] }),
                            new Paragraph({ 
                            	spacing: { after: 1000 },
                                children: [new TextRun({ text: "     "+clerk_name, bold: true, size: 30 })] }),
                          		 // Before the fix (Causes the error):
					// new Paragraph({
					//   tabStops: [{ type: TabStopType.RIGHT, position: 9000 }]
					// });

					// After the fix (Uses the global docx object):
					new Paragraph({
					    spacing: { after: 600 },
					    // Use docx.TabStopType
					    tabStops: [
						{
						    type: docx.TabStopType.RIGHT, // Access TabStopType through the global docx object
						    position: 9000, 
						},
					    ],
					    children: [
						new docx.TextRun({ // Also ensure Paragraph and TextRun are prefixed if necessary
						    text: clerk_desig, 
						    bold: true, 
						    size: 30 
						}),
						new docx.TextRun({ 
						    text: "\t", // The tab character
						}),
						new docx.TextRun({ 
						    text: checked_by, 
						    bold: true, 
						    size: 30 
						})
					    ] 
					}),
						
						
						
						
                            ...(index < data.length - 1 ? [new Paragraph({ children: [new PageBreak()] })] : []), // Add page break except for last row
                        ],
                    };
                });

                let doc = new Document({ sections });

                try {
                    const blob = await Packer.toBlob(doc);
                    saveAs(blob, "Converted_Document.docx");
                } catch (error) {
                    console.error("Error generating Word document:", error);
                    alert("Failed to generate document. See console for details.");
                }
            }


            document.getElementById("downloadBtn").addEventListener("click", function() {
            // Replace 'sample.xlsx' with the correct file path
            const fileUrl = "data.xlsx";
            
            // Create an invisible anchor element
            const link = document.createElement("a");
            link.href = fileUrl;
            link.download = "data.xlsx"; // Set the download attribute
            
            // Append the anchor to the body and trigger the click event
            document.body.appendChild(link);
            link.click();
            
            // Remove the anchor element after download
            document.body.removeChild(link);
        });



    </script>
</body>
</html>
