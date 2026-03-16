document.getElementById('charter-form').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent the form from submitting normally

    // --- Helper Functions ---
    const createHeading = (text) => new docx.Paragraph({
        text: text,
        heading: docx.HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 }
    });

    const createSubheading = (text) => new docx.Paragraph({
        text: text,
        style: "strong",
        spacing: { after: 60 }
    });

    const createBullet = (text) => new docx.Paragraph({
        text: text,
        bullet: { level: 0 },
    });

    const createParagraph = (text) => new docx.Paragraph({
        text: text || "", // Ensure text is not undefined
    });

    // --- Get Form Data ---
    const getInputValue = (id) => document.getElementById(id).value;
    const getListFromTextarea = (id) => getInputValue(id).split('\n').filter(line => line.trim() !== '');

    const charterName = getInputValue('charter-name');
    const vision = getInputValue('vision-mission');
    const purpose = getInputValue('purpose');
    const objectives = getListFromTextarea('objectives');
    const metrics = getListFromTextarea('success-metrics');
    const inScope = getListFromTextarea('in-scope');
    const outOfScope = getListFromTextarea('out-of-scope');
    const chair = getInputValue('chair-lead');
    const membersText = getListFromTextarea('members');
    const frequency = getInputValue('meeting-frequency');
    const decisionMaking = getInputValue('decision-making');
    const deliverables = getListFromTextarea('key-deliverables');
    const term = getInputValue('term');

    // --- Build Document Sections ---

    // **FIXED SECTION:** Safely create the member table rows
    const memberRows = membersText.map(memberLine => {
        const parts = memberLine.split(',');
        const name = parts[0] ? parts[0].trim() : ''; // Get the name, or empty string if not present
        const title = parts[1] ? parts[1].trim() : ''; // Get the title, or empty string if not present
        return new docx.TableRow({
            children: [
                new docx.TableCell({ children: [createParagraph(name)] }),
                new docx.TableCell({ children: [createParagraph(title)] }),
            ],
        });
    });

    const memberTable = new docx.Table({
        rows: [
            new docx.TableRow({
                children: [
                    new docx.TableCell({ children: [new docx.Paragraph({ text: "Name", style: "strong" })] }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: "Title/Role", style: "strong" })] }),
                ],
            }),
            ...memberRows
        ],
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
    });

    // Create the document
    const doc = new docx.Document({
        sections: [{
            children: [
                new docx.Paragraph({ text: charterName, heading: docx.HeadingLevel.TITLE, alignment: docx.AlignmentType.CENTER }),
                createHeading("1.0 Vision & Mission"),
                createSubheading("Vision"),
                createParagraph(vision),

                createHeading("2.0 Purpose & Goals"),
                createSubheading("Purpose"),
                createParagraph(purpose),
                createSubheading("Objectives"),
                ...objectives.map(createBullet),
                createSubheading("Success Metrics"),
                ...metrics.map(createBullet),

                createHeading("3.0 Scope"),
                createSubheading("In-Scope"),
                ...inScope.map(createBullet),
                createSubheading("Out-of-Scope"),
                ...outOfScope.map(createBullet),

                createHeading("4.0 Membership & Roles"),
                createSubheading("Chair/Lead"),
                createParagraph(chair),
                createSubheading("Membership"),
                memberTable,

                createHeading("5.0 Operations & Governance"),
                createSubheading("Meeting Frequency"),
                createParagraph(frequency),
                createSubheading("Decision-Making"),
                createParagraph(decisionMaking),
                createSubheading("Key Deliverables"),
                ...deliverables.map(createBullet),

                createHeading("6.0 Charter Review"),
                createSubheading("Term"),
                createParagraph(term),
            ],
        }],
    });

    // --- Generate and Download ---
    docx.Packer.toBlob(doc).then(blob => {
        // Use FileSaver.js to trigger the download
        saveAs(blob, `${charterName.replace(/\s+/g, '_') || "Charter"}_Charter.docx`);
    }).catch(err => {
        console.error(err); // Log any error that happens during document creation
    });
});
