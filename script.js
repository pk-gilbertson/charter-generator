document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('charter-form');
  const statusEl = document.getElementById('form-status');
  const fillButton = document.getElementById('fill-starter-content');
  const filenamePreview = document.getElementById('filename-preview');
  const completionText = document.getElementById('completion-text');
  const completionBar = document.getElementById('completion-bar');
  const submitButton = form?.querySelector('button[type="submit"]');
  const submitLabel = submitButton?.textContent?.trim() || 'Generate Charter (.docx)';
  function getDocx() { return window.docx; }

  if (!form) {
    console.error('Charter form not found.');
    return;
  }

  const DEFAULTS = {
    'agency-name': 'Agency Name',
    'charter-name': 'Data Governance Steering Committee',
    'committee-type': 'Data Governance Steering Committee',
    'agency-scope': 'Agency-wide',
    'executive-sponsor': 'Executive Sponsor',
    'chair-lead': 'Committee Chair',
    'effective-date': new Date().toISOString().slice(0, 10),
    'term-review': 'Effective until revised or rescinded; reviewed annually.',
    purpose:
      'The purpose of this committee is to establish direction, accountability, and oversight for the management and use of data as a strategic asset in support of agency operations, policy, reporting, and responsible innovation.',
    'vision-mission':
      'Vision: Trusted, timely, secure, and well-understood data supports better services, decision-making, and public stewardship.\n\nMission: To guide agency-wide data governance through clear roles, practical standards, coordinated decision-making, and responsible access and use.',
    objectives: [
      'Promote consistent accountability for priority data assets.',
      'Improve data quality, documentation, and standardization.',
      'Support lawful, secure, and efficient data sharing and access.',
      'Resolve cross-functional data issues and decision points.',
      'Advance a practical, sustainable culture of data governance.'
    ].join('\n'),
    'success-metrics': [
      'Priority data domains have assigned owners and stewards.',
      'Core definitions and standards are documented and approved.',
      'Data issues are tracked and resolved through a defined process.',
      'Requests for data access or sharing are reviewed consistently.',
      'Governance deliverables are completed according to committee priorities.'
    ].join('\n'),
    'in-scope': [
      'Data standards, definitions, and business rules.',
      'Data quality priorities and issue resolution.',
      'Metadata, documentation, and stewardship practices.',
      'Data access, sharing, and escalation workflows.',
      'Governance priorities related to reporting, analytics, and responsible data use.'
    ].join('\n'),
    'out-of-scope': [
      'Routine system administration and platform maintenance.',
      'Project management activities outside approved governance responsibilities.',
      'Operational decisions that remain within program management authority unless escalated.'
    ].join('\n'),
    'decision-authority':
      'The committee may approve governance standards, recommend policy changes, review and prioritize governance issues, and escalate decisions that require executive or enterprise-level action.',
    'escalation-path':
      'Issues that cannot be resolved by the committee, or that carry enterprise, legal, privacy, security, or significant operational impact, will be escalated through the executive sponsor and appropriate leadership channels.',
    'guiding-principles': [
      'Treat data as a strategic asset.',
      'Protect sensitive and regulated information.',
      'Promote responsible access and appropriate sharing.',
      'Standardize where practical while respecting business context.',
      'Assign clear accountability for data decisions.',
      'Use governance to enable operations, not create unnecessary burden.'
    ].join('\n'),
    members: [
      'Jane Doe, Chief Data Officer, Chair, Voting',
      'John Smith, Program Director, Member, Voting',
      'Mary Jones, Privacy Officer, Advisor, Non-Voting',
      'Alex Brown, IT Director, Member, Voting'
    ].join('\n'),
    'required-functions': [
      'Program or business leadership',
      'Information technology',
      'Privacy and security',
      'Legal or compliance',
      'Analytics, reporting, or performance management',
      'Records, finance, or other agency-specific perspectives as needed'
    ].join('\n'),
    'role-definitions': [
      'Executive Sponsor: Provides executive support, alignment, and escalation authority.',
      'Chair: Leads meetings, sets priorities, and guides committee decisions.',
      'Members: Participate in deliberation, decision-making, and follow-through.',
      'Advisors: Provide subject matter expertise in support of the committee.',
      'Data Owners, Stewards, and Custodians: Support governance decisions within their assigned roles.'
    ].join('\n'),
    responsibilities: [
      'Review and approve governance priorities, standards, and supporting guidance.',
      'Clarify ownership, stewardship, and accountability for priority data assets.',
      'Monitor governance issues, risks, and implementation progress.',
      'Resolve or escalate conflicts related to data definitions, quality, access, and use.',
      'Support practical coordination across business, technical, privacy, and legal stakeholders.'
    ].join('\n'),
    'annual-priorities': [
      'Establish a governance issue intake and tracking process.',
      'Document core data elements and definitions.',
      'Assign accountable roles for priority datasets.',
      'Create or refine standard templates for governance and sharing.'
    ].join('\n'),
    'key-deliverables': [
      'Committee charter',
      'Governance issue log',
      'Priority data glossary or data dictionary',
      'Standards, guidelines, or decision records',
      'Periodic status or progress summary'
    ].join('\n'),
    'meeting-frequency': 'Monthly',
    quorum: 'A simple majority of voting members.',
    'decision-making': 'Consensus where practical; otherwise majority vote.',
    'meeting-administration':
      'The chair or designee will prepare agendas, document decisions, maintain meeting records, and track action items and escalations.',
    'policy-alignment':
      'The committee will operate in alignment with applicable laws, regulations, statewide policy, agency policy, privacy requirements, security expectations, and records management obligations.',
    'privacy-security-considerations':
      'Privacy, security, legal, and other control functions will be engaged when governance issues involve confidential data, regulated data, release decisions, new uses of data, or elevated risk.',
    'data-sharing':
      'The committee may review or support processes related to internal sharing, external sharing, access requests, and associated agreements or approvals, consistent with agency and enterprise requirements.',
    subcommittees: 'Data Quality Working Group\nMetadata and Standards Working Group',
    'version-history': `1.0, ${new Date().toISOString().slice(0, 10)}, System, Initial charter generated`
  };

  const REQUIRED_FIELD_IDS = ['agency-name', 'charter-name'];
  const PROGRESS_FIELD_IDS = [
    'agency-name',
    'charter-name',
    'purpose',
    'vision-mission',
    'in-scope',
    'guiding-principles',
    'members',
    'responsibilities',
    'meeting-frequency'
  ];

  function getField(id) {
    return document.getElementById(id);
  }

  function getOptionalValue(id) {
    const field = getField(id);
    if (!field || typeof field.value !== 'string') return '';
    return field.value.trim();
  }

  function getValue(id, fallback = '') {
    const value = getOptionalValue(id);
    return value || fallback;
  }

  function setValue(id, value) {
    const field = getField(id);
    if (!field) return;
    field.value = value;
  }

  function setStatus(message = '', state = '') {
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.classList.remove('success', 'error');
    if (state) statusEl.classList.add(state);
  }

  function safeFileName(text) {
    const cleaned = String(text || 'Data_Governance_Charter')
      .trim()
      .replace(/[^a-z0-9]+/gi, '_')
      .replace(/^_+|_+$/g, '');
    return cleaned || 'Data_Governance_Charter';
  }

  function toLines(text) {
    return String(text || '')
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function splitWithLimit(line, expectedParts) {
    const rawParts = String(line || '')
      .split(',')
      .map((part) => part.trim());

    if (rawParts.length <= expectedParts) {
      while (rawParts.length < expectedParts) rawParts.push('');
      return rawParts;
    }

    const fixed = rawParts.slice(0, expectedParts - 1);
    fixed.push(rawParts.slice(expectedParts - 1).join(', ').trim());
    return fixed;
  }

  function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(`${dateString}T00:00:00`);
    if (Number.isNaN(date.getTime())) return dateString;

    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
  }

  function setAriaInvalid(id, invalid) {
    const field = getField(id);
    if (!field) return;
    if (invalid) {
      field.setAttribute('aria-invalid', 'true');
    } else {
      field.removeAttribute('aria-invalid');
    }
  }

  function validateForm() {
    let firstInvalid = null;

    REQUIRED_FIELD_IDS.forEach((id) => {
      const invalid = !getOptionalValue(id);
      setAriaInvalid(id, invalid);
      if (invalid && !firstInvalid) firstInvalid = getField(id);
    });

    return {
      isValid: !firstInvalid,
      firstInvalid
    };
  }

  function updateFilenamePreview() {
    if (!filenamePreview) return;
    const charterName = getValue('charter-name', DEFAULTS['charter-name']);
    filenamePreview.textContent = `${safeFileName(charterName)}_Charter.docx`;
  }

  function updateCompletion() {
    if (!completionText || !completionBar) return;

    const completed = PROGRESS_FIELD_IDS.filter((id) => getOptionalValue(id)).length;
    const percent = Math.round((completed / PROGRESS_FIELD_IDS.length) * 100);

    completionBar.style.width = `${percent}%`;

    if (percent < 35) {
      completionText.textContent = 'Start with the charter basics and purpose sections.';
    } else if (percent < 70) {
      completionText.textContent = 'The draft is taking shape. Add membership, responsibilities, and operating details next.';
    } else if (percent < 100) {
      completionText.textContent = 'Nearly complete. Review advanced sections and export when ready.';
    } else {
      completionText.textContent = 'Core sections are complete. You are ready to generate the charter.';
    }
  }

  function updateHelpers() {
    updateFilenamePreview();
    updateCompletion();
  }

  function fillStarterContent() {
    Object.entries(DEFAULTS).forEach(([id, value]) => {
      const field = getField(id);
      if (!field) return;
      if (!getOptionalValue(id)) {
        setValue(id, value);
      }
    });

    updateHelpers();
    setStatus('Starter content loaded. Review and customize before export.', 'success');
  }

  function blankParagraph(after = 120) {
    return new getDocx().Paragraph({
      text: '',
      spacing: { after }
    });
  }

  function heading(text, level = getDocx().HeadingLevel.HEADING_1, color = '1F2933') {
    return new getDocx().Paragraph({
      heading: level,
      spacing: { before: 220, after: 80 },
      children: [
        new getDocx().TextRun({
          text,
          bold: true,
          color
        })
      ]
    });
  }

  function bodyParagraph(text) {
    return new getDocx().Paragraph({
      children: [new getDocx().TextRun({ text })],
      spacing: { after: 120 }
    });
  }

  function multiParagraphs(text) {
    return toLines(text).map((line) => bodyParagraph(line));
  }

  function bulletList(text) {
    return toLines(text).map(
      (item) =>
        new getDocx().Paragraph({
          text: item,
          bullet: { level: 0 },
          spacing: { after: 80 }
        })
    );
  }

  function tableCell(text, options = {}) {
    return new getDocx().TableCell({
      shading: options.shading ? { fill: options.shading } : undefined,
      verticalAlign: getDocx().VerticalAlign.CENTER,
      children: [
        new getDocx().Paragraph({
          spacing: { before: 60, after: 60 },
          children: [
            new getDocx().TextRun({
              text: String(text || ''),
              bold: Boolean(options.bold),
              color: options.color || '1F2933'
            })
          ]
        })
      ]
    });
  }

  function createKeyValueTable(rows) {
    return new getDocx().Table({
      width: { size: 100, type: getDocx().WidthType.PERCENTAGE },
      layout: getDocx().TableLayoutType.FIXED,
      borders: {
        top: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        bottom: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        left: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        right: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        insideHorizontal: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'E8DDD0' },
        insideVertical: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'E8DDD0' }
      },
      rows: rows.map(
        ([label, value]) =>
          new getDocx().TableRow({
            children: [
              new getDocx().TableCell({
                width: { size: 28, type: getDocx().WidthType.PERCENTAGE },
                shading: { fill: 'F7F3EC' },
                verticalAlign: getDocx().VerticalAlign.CENTER,
                children: [
                  new getDocx().Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new getDocx().TextRun({ text: label, bold: true, color: '1F2933' })]
                  })
                ]
              }),
              new getDocx().TableCell({
                width: { size: 72, type: getDocx().WidthType.PERCENTAGE },
                verticalAlign: getDocx().VerticalAlign.CENTER,
                children: [
                  new getDocx().Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new getDocx().TextRun({ text: String(value || '') })]
                  })
                ]
              })
            ]
          })
      )
    });
  }

  function createDataTable(headers, linesText, expectedParts, fallbackText, headerFill, headerTextColor) {
    const lines = toLines(linesText || fallbackText);
    const rows = lines.map((line) => splitWithLimit(line, expectedParts));

    return new getDocx().Table({
      width: { size: 100, type: getDocx().WidthType.PERCENTAGE },
      layout: getDocx().TableLayoutType.FIXED,
      borders: {
        top: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        bottom: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        left: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        right: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        insideHorizontal: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'E8DDD0' },
        insideVertical: { style: getDocx().BorderStyle.SINGLE, size: 1, color: 'E8DDD0' }
      },
      rows: [
        new getDocx().TableRow({
          tableHeader: true,
          children: headers.map((header) =>
            tableCell(header, {
              bold: true,
              shading: headerFill,
              color: headerTextColor
            })
          )
        }),
        ...rows.map(
          (row, index) =>
            new getDocx().TableRow({
              children: row.map((value) =>
                tableCell(value, {
                  shading: index % 2 === 0 ? 'FFFFFF' : 'FBF8F2'
                })
              )
            })
        )
      ]
    });
  }

  function createSection(title, color, buildContent) {
    const content = buildContent();
    if (!Array.isArray(content) || content.length === 0) return [];
    return [heading(title, getDocx().HeadingLevel.HEADING_1, color), ...content, blankParagraph()];
  }

  function buildDocument() {
    const charterName = getValue('charter-name', DEFAULTS['charter-name']);
    const agencyName = getValue('agency-name', DEFAULTS['agency-name']);
    const committeeType = getValue('committee-type', DEFAULTS['committee-type']);
    const agencyScope = getValue('agency-scope', DEFAULTS['agency-scope']);
    const executiveSponsor = getValue('executive-sponsor', DEFAULTS['executive-sponsor']);
    const chairLead = getValue('chair-lead', DEFAULTS['chair-lead']);
    const effectiveDate = formatDate(getValue('effective-date', DEFAULTS['effective-date']));
    const termReview = getValue('term-review', DEFAULTS['term-review']);

    const metadataTable = createKeyValueTable([
      ['Agency / Department', agencyName],
      ['Charter Name', charterName],
      ['Committee Type', committeeType],
      ['Organizational Scope', agencyScope],
      ['Executive Sponsor', executiveSponsor],
      ['Chair / Lead', chairLead],
      ['Effective Date', effectiveDate],
      ['Term & Review Cycle', termReview]
    ]);

    const membersTable = createDataTable(
      ['Name', 'Title', 'Role', 'Voting Status'],
      getValue('members', DEFAULTS.members),
      4,
      DEFAULTS.members,
      '2F5D50',
      'FFFFFF'
    );

    const versionHistoryTable = createDataTable(
      ['Version', 'Date', 'Author', 'Summary of Changes'],
      getOptionalValue('version-history') || DEFAULTS['version-history'],
      4,
      DEFAULTS['version-history'],
      '5A2D5C',
      'FFFFFF'
    );

    const children = [
      new getDocx().Paragraph({
        alignment: getDocx().AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new getDocx().TextRun({ text: charterName, bold: true, size: 34, color: '1F2933' })]
      }),
      new getDocx().Paragraph({
        alignment: getDocx().AlignmentType.CENTER,
        spacing: { after: 220 },
        children: [new getDocx().TextRun({ text: agencyName, size: 24, color: '5A6470' })]
      }),
      metadataTable,
      blankParagraph(60),

      ...createSection('1. Purpose', 'A54A2A', () => multiParagraphs(getValue('purpose', DEFAULTS.purpose))),
      ...createSection('2. Vision / Mission', 'A54A2A', () => multiParagraphs(getValue('vision-mission', DEFAULTS['vision-mission']))),
      ...createSection('3. Objectives', 'A54A2A', () => bulletList(getValue('objectives', DEFAULTS.objectives))),
      ...createSection('4. Success Metrics', 'A54A2A', () => bulletList(getValue('success-metrics', DEFAULTS['success-metrics']))),

      ...createSection('5. Scope & Authority', 'A54A2A', () => [
        heading('In Scope', getDocx().HeadingLevel.HEADING_2, 'A54A2A'),
        ...bulletList(getValue('in-scope', DEFAULTS['in-scope'])),
        heading('Out of Scope', getDocx().HeadingLevel.HEADING_2, 'A54A2A'),
        ...bulletList(getValue('out-of-scope', DEFAULTS['out-of-scope'])),
        heading('Decision Authority', getDocx().HeadingLevel.HEADING_2, 'A54A2A'),
        ...multiParagraphs(getValue('decision-authority', DEFAULTS['decision-authority'])),
        heading('Escalation Path', getDocx().HeadingLevel.HEADING_2, 'A54A2A'),
        ...multiParagraphs(getValue('escalation-path', DEFAULTS['escalation-path']))
      ]),

      ...createSection('6. Guiding Principles', '2F5D50', () =>
        bulletList(getValue('guiding-principles', DEFAULTS['guiding-principles']))
      ),

      ...createSection('7. Membership & Representation', '2F5D50', () => [
        heading('Committee Members', getDocx().HeadingLevel.HEADING_2, '2F5D50'),
        membersTable,
        heading('Required Functions / Perspectives', getDocx().HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('required-functions', DEFAULTS['required-functions'])),
        heading('Role Definitions', getDocx().HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('role-definitions', DEFAULTS['role-definitions']))
      ]),

      ...createSection('8. Responsibilities & Deliverables', '2F5D50', () => [
        heading('Committee Responsibilities', getDocx().HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('responsibilities', DEFAULTS.responsibilities)),
        heading('Annual or Initial Priorities', getDocx().HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('annual-priorities', DEFAULTS['annual-priorities'])),
        heading('Key Deliverables', getDocx().HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('key-deliverables', DEFAULTS['key-deliverables']))
      ]),

      ...createSection('9. Operating Model', '5A2D5C', () => {
        const operatingModelTable = createKeyValueTable([
          ['Meeting Cadence', getValue('meeting-frequency', DEFAULTS['meeting-frequency'])],
          ['Quorum', getValue('quorum', DEFAULTS.quorum)],
          ['Decision-Making Process', getValue('decision-making', DEFAULTS['decision-making'])]
        ]);

        return [
          operatingModelTable,
          heading('Meeting Administration', getDocx().HeadingLevel.HEADING_2, '5A2D5C'),
          ...multiParagraphs(getValue('meeting-administration', DEFAULTS['meeting-administration']))
        ];
      }),

      ...createSection('10. Policy, Privacy, Security & Sharing', '5A2D5C', () => {
        const content = [];
        const policyAlignment = getOptionalValue('policy-alignment');
        const privacySecurity = getOptionalValue('privacy-security-considerations');
        const dataSharing = getOptionalValue('data-sharing');

        if (policyAlignment) {
          content.push(heading('Policy / Legal / Regulatory Alignment', getDocx().HeadingLevel.HEADING_2, '5A2D5C'));
          content.push(...multiParagraphs(policyAlignment));
        }
        if (privacySecurity) {
          content.push(heading('Privacy, Security & Data Release Considerations', getDocx().HeadingLevel.HEADING_2, '5A2D5C'));
          content.push(...multiParagraphs(privacySecurity));
        }
        if (dataSharing) {
          content.push(heading('Data Sharing & Access Considerations', getDocx().HeadingLevel.HEADING_2, '5A2D5C'));
          content.push(...multiParagraphs(dataSharing));
        }

        return content;
      }),

      ...createSection('11. Working Groups & Subcommittees', '5A2D5C', () => {
        const subcommittees = getOptionalValue('subcommittees');
        return subcommittees ? bulletList(subcommittees) : [];
      }),

      ...createSection('12. Version History', '5A2D5C', () => [versionHistoryTable])
    ];

    return new getDocx().Document({
      creator: 'Data Governance Charter Generator',
      title: charterName,
      description: 'Generated charter document for a data governance committee.',
      sections: [{ children }]
    });
  }

  form.addEventListener('input', (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;

    if (target.id && REQUIRED_FIELD_IDS.includes(target.id)) {
      setAriaInvalid(target.id, !getOptionalValue(target.id));
    }

    updateHelpers();
  });

  form.addEventListener('reset', () => {
    window.setTimeout(() => {
      REQUIRED_FIELD_IDS.forEach((id) => setAriaInvalid(id, false));
      setStatus('');
      updateHelpers();
    }, 0);
  });

  if (fillButton) {
    fillButton.addEventListener('click', fillStarterContent);
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const { isValid, firstInvalid } = validateForm();
    if (!isValid) {
      setStatus('Please complete the required charter basics before generating the document.', 'error');
      firstInvalid?.focus();
      return;
    }

    if (!getDocx() || typeof window.saveAs !== 'function') {
      setStatus('Document libraries did not load. Refresh the page and try again.', 'error');
      return;
    }

    try {
      if (submitButton) {
        submitButton.disabled = true;
        submitButton.textContent = 'Generating...';
      }

      setStatus('Generating your charter document...', 'success');

      const documentDefinition = buildDocument();
      const blob = await getDocx().Packer.toBlob(documentDefinition);
      const fileName = `${safeFileName(getValue('charter-name', DEFAULTS['charter-name']))}_Charter.docx`;

      window.saveAs(blob, fileName);
      setStatus(`Done. Download started for ${fileName}.`, 'success');
    } catch (error) {
      console.error('Error generating charter:', error);
      setStatus('There was an error generating the charter. Check the browser console for details.', 'error');
    } finally {
      if (submitButton) {
        submitButton.disabled = false;
        submitButton.textContent = submitLabel;
      }
    }
  });

  updateHelpers();
});
