document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('charter-form');
  const statusEl = document.getElementById('form-status');
  const fillButton = document.getElementById('fill-starter-content');
  const filenamePreview = document.getElementById('filename-preview');
  const completionText = document.getElementById('completion-text');
  const completionBar = document.getElementById('completion-bar');
  const submitButton = form?.querySelector('button[type="submit"]');
  const submitLabel = submitButton?.textContent?.trim() || 'Generate Charter (.docx)';
  const membersTbody = document.getElementById('members-tbody');
  const addMemberBtn = document.getElementById('add-member-row');
  const jumpToTopButton = document.getElementById('jump-to-top');

  let docx = null;

  if (!form) {
    console.error('Charter form not found.');
    return;
  }

  const MEMBER_COLUMNS = ['Name', 'Title', 'Role', 'Voting Status'];
  const VOTING_OPTIONS = ['', 'Voting', 'Non-Voting'];

  const DEFAULT_MEMBERS = [
    { name: 'Jane Doe', title: 'Chief Data Officer', role: 'Chair', voting: 'Voting' },
    { name: 'John Smith', title: 'Program Director', role: 'Member', voting: 'Voting' },
    { name: 'Mary Jones', title: 'Privacy Officer', role: 'Advisor', voting: 'Non-Voting' },
    { name: 'Alex Brown', title: 'IT Director', role: 'Member', voting: 'Voting' }
  ];

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
    'responsibilities',
    'meeting-frequency'
  ];

  function resolveDocx() {
    const library = window.docx;
    if (!library) return null;

    const requiredKeys = [
      'Document',
      'Paragraph',
      'TextRun',
      'Table',
      'TableRow',
      'TableCell',
      'HeadingLevel',
      'AlignmentType',
      'WidthType',
      'TableLayoutType',
      'VerticalAlign',
      'BorderStyle',
      'Packer'
    ];

    return requiredKeys.every((key) => key in library) ? library : null;
  }

  function createMemberRow(data = {}) {
    const tr = document.createElement('tr');
    tr.className = 'members-row';

    const textFields = [
      { key: 'name', placeholder: 'e.g., Jane Doe', label: 'Name' },
      { key: 'title', placeholder: 'e.g., Chief Data Officer', label: 'Title' },
      { key: 'role', placeholder: 'e.g., Chair', label: 'Role' }
    ];

    textFields.forEach(({ key, placeholder, label }) => {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'members-input';
      input.placeholder = placeholder;
      input.value = data[key] || '';
      input.setAttribute('aria-label', label);
      input.addEventListener('input', updateHelpers);
      td.appendChild(input);
      tr.appendChild(td);
    });

    const tdVoting = document.createElement('td');
    const select = document.createElement('select');
    select.className = 'members-input';
    select.setAttribute('aria-label', 'Voting Status');

    VOTING_OPTIONS.forEach((optionValue) => {
      const option = document.createElement('option');
      option.value = optionValue;
      option.textContent = optionValue || 'Select status';
      select.appendChild(option);
    });

    select.value = data.voting || '';
    select.addEventListener('change', updateHelpers);
    tdVoting.appendChild(select);
    tr.appendChild(tdVoting);

    const tdDel = document.createElement('td');
    tdDel.className = 'members-cell--delete';

    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'members-delete-btn';
    delBtn.setAttribute('aria-label', 'Remove this member');
    delBtn.title = 'Remove row';
    delBtn.innerHTML = '&times;';
    delBtn.addEventListener('click', () => {
      tr.remove();
      if (!membersTbody?.querySelector('tr.members-row')) {
        initMembersTable();
      }
      updateHelpers();
    });

    tdDel.appendChild(delBtn);
    tr.appendChild(tdDel);

    return tr;
  }

  function initMembersTable() {
    if (!membersTbody) return;
    membersTbody.innerHTML = '';
    membersTbody.appendChild(createMemberRow());
  }

  function populateMembersTable(members) {
    if (!membersTbody) return;
    membersTbody.innerHTML = '';
    members.forEach((member) => membersTbody.appendChild(createMemberRow(member)));
  }

  function clearMembersTable() {
    initMembersTable();
  }

  function getMemberRows() {
    if (!membersTbody) return [];

    return Array.from(membersTbody.querySelectorAll('tr.members-row')).map((tr) => {
      const fields = tr.querySelectorAll('.members-input');
      return {
        name: fields[0]?.value.trim() || '',
        title: fields[1]?.value.trim() || '',
        role: fields[2]?.value.trim() || '',
        voting: fields[3]?.value.trim() || ''
      };
    });
  }

  function hasMemberData() {
    return getMemberRows().some((row) => row.name || row.title || row.role || row.voting);
  }

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

    return { isValid: !firstInvalid, firstInvalid };
  }

  function updateFilenamePreview() {
    if (!filenamePreview) return;
    const charterName = getValue('charter-name', DEFAULTS['charter-name']);
    filenamePreview.textContent = `${safeFileName(charterName)}_Charter.docx`;
  }

  function updateCompletion() {
    if (!completionText || !completionBar) return;

    const textCompleted = PROGRESS_FIELD_IDS.filter((id) => getOptionalValue(id)).length;
    const membersCompleted = hasMemberData() ? 1 : 0;
    const total = PROGRESS_FIELD_IDS.length + 1;
    const completed = textCompleted + membersCompleted;
    const percent = Math.round((completed / total) * 100);

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

  function updateJumpToTopVisibility() {
    if (!jumpToTopButton) return;

    const triggerPoint = Math.max(420, window.innerHeight * 0.6);
    const shouldShow = window.scrollY > triggerPoint;
    jumpToTopButton.classList.toggle('is-visible', shouldShow);
  }

  function handleJumpToTop() {
    window.scrollTo({
      top: 0,
      behavior: 'smooth'
    });
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

    if (!hasMemberData()) {
      populateMembersTable(DEFAULT_MEMBERS);
    }

    updateHelpers();
    setStatus('Starter content loaded. Review and customize before export.', 'success');
  }

  function blankParagraph(after = 120) {
    return new docx.Paragraph({ text: '', spacing: { after } });
  }

  function heading(text, level, color = '1F2933') {
    return new docx.Paragraph({
      heading: level,
      spacing: { before: 220, after: 80 },
      children: [new docx.TextRun({ text, bold: true, color })]
    });
  }

  function bodyParagraph(text) {
    return new docx.Paragraph({
      children: [new docx.TextRun({ text })],
      spacing: { after: 120 }
    });
  }

  function multiParagraphs(text) {
    return toLines(text).map((line) => bodyParagraph(line));
  }

  function bulletList(text) {
    return toLines(text).map(
      (item) => new docx.Paragraph({ text: item, bullet: { level: 0 }, spacing: { after: 80 } })
    );
  }

  function tableCell(text, options = {}) {
    return new docx.TableCell({
      shading: options.shading ? { fill: options.shading } : undefined,
      verticalAlign: docx.VerticalAlign.CENTER,
      children: [
        new docx.Paragraph({
          spacing: { before: 60, after: 60 },
          children: [
            new docx.TextRun({
              text: String(text || ''),
              bold: Boolean(options.bold),
              color: options.color || '1F2933'
            })
          ]
        })
      ]
    });
  }

  function getTableBorders() {
    return {
      top: { style: docx.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
      bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
      left: { style: docx.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
      right: { style: docx.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
      insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: 'E8DDD0' },
      insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: 'E8DDD0' }
    };
  }

  function createKeyValueTable(rows) {
    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(),
      rows: rows.map(
        ([label, value]) =>
          new docx.TableRow({
            children: [
              new docx.TableCell({
                width: { size: 28, type: docx.WidthType.PERCENTAGE },
                shading: { fill: 'F7F3EC' },
                verticalAlign: docx.VerticalAlign.CENTER,
                children: [
                  new docx.Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new docx.TextRun({ text: label, bold: true, color: '1F2933' })]
                  })
                ]
              }),
              new docx.TableCell({
                width: { size: 72, type: docx.WidthType.PERCENTAGE },
                verticalAlign: docx.VerticalAlign.CENTER,
                children: [
                  new docx.Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new docx.TextRun({ text: String(value || '') })]
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

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(),
      rows: [
        new docx.TableRow({
          tableHeader: true,
          children: headers.map((header) =>
            tableCell(header, { bold: true, shading: headerFill, color: headerTextColor })
          )
        }),
        ...rows.map((row, index) =>
          new docx.TableRow({
            children: row.map((value) =>
              tableCell(value, { shading: index % 2 === 0 ? 'FFFFFF' : 'FBF8F2' })
            )
          })
        )
      ]
    });
  }

  function createMembersDocTable(members) {
    const rows = members.length > 0 ? members : [{ name: '', title: '', role: '', voting: '' }];

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      layout: docx.TableLayoutType.FIXED,
      borders: getTableBorders(),
      rows: [
        new docx.TableRow({
          tableHeader: true,
          children: MEMBER_COLUMNS.map((col) =>
            tableCell(col, { bold: true, shading: '2F5D50', color: 'FFFFFF' })
          )
        }),
        ...rows.map((member, index) =>
          new docx.TableRow({
            children: [member.name, member.title, member.role, member.voting].map((value) =>
              tableCell(value, { shading: index % 2 === 0 ? 'FFFFFF' : 'FBF8F2' })
            )
          })
        )
      ]
    });
  }

  function createSection(title, color, buildContent) {
    const content = buildContent();
    if (!Array.isArray(content) || content.length === 0) return [];
    return [heading(title, docx.HeadingLevel.HEADING_1, color), ...content, blankParagraph()];
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

    const memberRows = getMemberRows().filter((row) => row.name || row.title || row.role || row.voting);
    const votingMembers = memberRows.filter((row) => row.voting === 'Voting');
    const nonVotingMembers = memberRows.filter((row) => row.voting === 'Non-Voting');

    const votingMembersTable = createMembersDocTable(votingMembers);
    const nonVotingMembersTable = createMembersDocTable(nonVotingMembers);

    const versionHistoryTable = createDataTable(
      ['Version', 'Date', 'Author', 'Summary of Changes'],
      getOptionalValue('version-history') || DEFAULTS['version-history'],
      4,
      DEFAULTS['version-history'],
      '5A2D5C',
      'FFFFFF'
    );

    const children = [
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new docx.TextRun({ text: charterName, bold: true, size: 34, color: '1F2933' })]
      }),
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 220 },
        children: [new docx.TextRun({ text: agencyName, size: 24, color: '5A6470' })]
      }),
      metadataTable,
      blankParagraph(60),

      ...createSection('1. Purpose', 'A54A2A', () => multiParagraphs(getValue('purpose', DEFAULTS.purpose))),
      ...createSection('2. Vision / Mission', 'A54A2A', () => multiParagraphs(getValue('vision-mission', DEFAULTS['vision-mission']))),
      ...createSection('3. Objectives', 'A54A2A', () => bulletList(getValue('objectives', DEFAULTS.objectives))),
      ...createSection('4. Success Metrics', 'A54A2A', () => bulletList(getValue('success-metrics', DEFAULTS['success-metrics']))),

      ...createSection('5. Scope & Authority', 'A54A2A', () => [
        heading('In Scope', docx.HeadingLevel.HEADING_2, 'A54A2A'),
        ...bulletList(getValue('in-scope', DEFAULTS['in-scope'])),
        heading('Out of Scope', docx.HeadingLevel.HEADING_2, 'A54A2A'),
        ...bulletList(getValue('out-of-scope', DEFAULTS['out-of-scope'])),
        heading('Decision Authority', docx.HeadingLevel.HEADING_2, 'A54A2A'),
        ...multiParagraphs(getValue('decision-authority', DEFAULTS['decision-authority'])),
        heading('Escalation Path', docx.HeadingLevel.HEADING_2, 'A54A2A'),
        ...multiParagraphs(getValue('escalation-path', DEFAULTS['escalation-path']))
      ]),

      ...createSection('6. Guiding Principles', '2F5D50', () =>
        bulletList(getValue('guiding-principles', DEFAULTS['guiding-principles']))
      ),

      ...createSection('7. Membership & Representation', '2F5D50', () => [
        heading('Committee Members - Voting', docx.HeadingLevel.HEADING_2, '2F5D50'),
        votingMembersTable,
        heading('Committee Members - Non-Voting', docx.HeadingLevel.HEADING_2, '2F5D50'),
        nonVotingMembersTable,
        heading('Required Functions / Perspectives', docx.HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('required-functions', DEFAULTS['required-functions'])),
        heading('Role Definitions', docx.HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('role-definitions', DEFAULTS['role-definitions']))
      ]),

      ...createSection('8. Responsibilities & Deliverables', '2F5D50', () => [
        heading('Committee Responsibilities', docx.HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('responsibilities', DEFAULTS.responsibilities)),
        heading('Annual or Initial Priorities', docx.HeadingLevel.HEADING_2, '2F5D50'),
        ...bulletList(getValue('annual-priorities', DEFAULTS['annual-priorities'])),
        heading('Key Deliverables', docx.HeadingLevel.HEADING_2, '2F5D50'),
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
          heading('Meeting Administration', docx.HeadingLevel.HEADING_2, '5A2D5C'),
          ...multiParagraphs(getValue('meeting-administration', DEFAULTS['meeting-administration']))
        ];
      }),

      ...createSection('10. Policy, Privacy, Security & Sharing', '5A2D5C', () => {
        const content = [];
        const policyAlignment = getOptionalValue('policy-alignment');
        const privacySecurity = getOptionalValue('privacy-security-considerations');
        const dataSharing = getOptionalValue('data-sharing');

        if (policyAlignment) {
          content.push(heading('Policy / Legal / Regulatory Alignment', docx.HeadingLevel.HEADING_2, '5A2D5C'));
          content.push(...multiParagraphs(policyAlignment));
        }

        if (privacySecurity) {
          content.push(heading('Privacy, Security & Data Release Considerations', docx.HeadingLevel.HEADING_2, '5A2D5C'));
          content.push(...multiParagraphs(privacySecurity));
        }

        if (dataSharing) {
          content.push(heading('Data Sharing & Access Considerations', docx.HeadingLevel.HEADING_2, '5A2D5C'));
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

    return new docx.Document({
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
      clearMembersTable();
      updateHelpers();
    }, 0);
  });

  if (fillButton) {
    fillButton.addEventListener('click', fillStarterContent);
  }

  if (addMemberBtn) {
    addMemberBtn.addEventListener('click', () => {
      if (!membersTbody) return;
      const newRow = createMemberRow();
      membersTbody.appendChild(newRow);
      newRow.querySelector('input, select')?.focus();
      updateHelpers();
    });
  }

  if (jumpToTopButton) {
    jumpToTopButton.addEventListener('click', handleJumpToTop);
  }

  window.addEventListener('scroll', updateJumpToTopVisibility, { passive: true });
  window.addEventListener('resize', updateJumpToTopVisibility);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const { isValid, firstInvalid } = validateForm();
    if (!isValid) {
      setStatus('Please complete the required charter basics before generating the document.', 'error');
      firstInvalid?.focus();
      return;
    }

    docx = resolveDocx();

    if (!docx || typeof window.saveAs !== 'function') {
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
      const blob = await docx.Packer.toBlob(documentDefinition);
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

  initMembersTable();
  updateHelpers();
  updateJumpToTopVisibility();
});
