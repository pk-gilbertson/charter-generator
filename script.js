document.getElementById('charter-form').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent the form from submitting normally

    // Helper function to create styled paragraphs
    const createHeading = (text) => new docx.Paragraph({
        text: text,
        heading: docx.HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 }
    });
    
    const createSubheading = (text) => new docx.Paragraph({
        text: text,
        style: "strong",
    });

    const createBullet = (text) => new docx.Paragraph({
        text: text,
        bullet: { level: 0 },
    });

    // Get form data
    const charterName = document.getElementById('charter-name').value;
    const vision = document.getElementById('vision-mission').value;
    const purpose = document.getElementById('purpose').value;
    const objectives = document.getElementById('objectives').value.split('\n').filter(line => line.trim() !== '');
    const metrics = document.getElementById('success-metrics').value.split('\n').filter(line => line.trim() !== '');
    const inScope = document.getElementById('in-scope').value.split('\n').filter(line => line.trim() !== '');
    const outOfScope = document.getElementById('out-of-scope').value.split('\n').filter(line => line.trim() !== '');
    const chair = document.getElementById('chair-lead').value;
    const members = document.getElementById('members').value.split('\n').filter(line => line.trim() !== '');
    const frequency = document.getElementById('meeting-frequency').value;
    const decisionMaking = document.getElementById('decision-making').value;
    const deliverables = document.getElementById('key-deliverables').value.split('\n').filter(line => line.trim() !== '');
    const term = document.getElementById('term').value;

    // Create member table rows
    const memberRows = members.map(member => {
        const [name, title] = member.split(',').map(s => s.trim());
        return new docx.TableRow({
            children: [
                new docx.TableCell({ children: [new docx.Paragraph(name || '')] }),
                new docx.TableCell({ children: [new docx.Paragraph(title || '')] }),
            ],
        });
    });

    // Create the document
    const doc = new docx.Document({
        sections: [{
            children: [
                new docx.Paragraph({
                    text: charterName,
                    heading: docx.HeadingLevel.TITLE,
                    alignment: docx.AlignmentType.CENTER,
                }),

                createHeading("1.0 Vision & Mission"),
                createSubheading("Vision"),
                new docx.Paragraph(vision),

                createHeading("2.0 Purpose & Goals"),
                createSubheading("Purpose"),
                new docx.Paragraph(purpose),
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
                new docx.Paragraph(chair),
                createSubheading("Membership"),
                new docx.Table({
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
                }),

                createHeading("5.0 Operations & Governance"),
                createSubheading("Meeting Frequency"),
                new docx.Paragraph(frequency),
                createSubheading("Decision-Making"),
                new docx.Paragraph(decisionMaking),
                createSubheading("Key Deliverables"),
                ...deliverables.map(createBullet),

                createHeading("6.0 Charter Review"),
                createSubheading("Term"),
                new docx.Paragraph(term),
            ],
        }],
    });

    // Generate and download the DOCX file
    docx.Packer.toBlob(doc).then(blob => {
        saveAs(blob, `${charterName.replace(/\s+/g, '_')}_Charter.docx`);
        console.log('Document created successfully');
    });
});
