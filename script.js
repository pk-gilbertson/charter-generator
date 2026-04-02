document.addEventListener('DOMContentLoaded', () => {
  const docxLib = window.docx;
  const form = document.getElementById('charter-form');
  const statusEl = document.getElementById('form-status');
  const submitButton = form?.querySelector('button[type="submit"]');
  const fillButton = document.getElementById('fill-starter-content');
  const filenamePreview = document.getElementById('filename-preview');
  const completionText = document.getElementById('completion-text');
  const completionBar = document.getElementById('completion-bar');
  const submitLabel = submitButton?.textContent || 'Generate Charter (.docx)';
  const committeeTypeSelect = document.getElementById('committee-type');
  const committeeTypeTrigger = document.getElementById('committee-type-trigger');
  const committeeTypeTriggerText = document.getElementById('committee-type-trigger-text');
  const committeeTypeMenu = document.getElementById('committee-type-menu');
  const committeeTypeOptions = document.getElementById('committee-type-options');
  const committeeTypePreviewText = document.getElementById('committee-type-preview-text');

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

  const COMMITTEE_TYPE_DEFINITIONS = {
    'Data Governance Steering Committee': 'Formal body with oversight and decision-making authority for governance.',
    'Data Governance Advisory Group': 'Provides guidance and recommendations, but does not usually hold final authority.',
    'Data Stewardship Council': 'Focuses on operational governance, stewardship, data quality, and standards.',
    'Working Group': 'Temporary or task-focused group addressing a specific governance need.',
    'Cross-Agency Governance Group': 'Coordinates governance across multiple agencies or departments.',
    Other: 'Use when the group follows a different or custom governance model.'
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

  function setStatus(message, state = '') {
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.classList.remove('success', 'error');
    if (state) statusEl.classList.add(state);
  }

  function getElement(id) {
    return document.getElementById(id);
  }

  function getValue(id, fallback = '') {
    const el = getElement(id);
    if (!el) return fallback;
    const value = String(el.value || '').trim();
    return value || fallback;
  }

  function getOptionalValue(id) {
    const el = getElement(id);
    return el ? String(el.value || '').trim() : '';
  }

  function setValue(id, value) {
    const el = getElement(id);
    if (!el) return;
    el.value = value;
    if (id === 'committee-type') {
      syncCommitteeTypePicker();
    }
  }

  function getCommitteeTypeDefinition(value) {
    return COMMITTEE_TYPE_DEFINITIONS[value] || COMMITTEE_TYPE_DEFINITIONS.Other;
  }

  function updateCommitteeTypePreview(value) {
    if (!committeeTypePreviewText) return;
    committeeTypePreviewText.textContent = getCommitteeTypeDefinition(value);
  }

  function syncCommitteeTypePicker() {
    if (!committeeTypeSelect || !committeeTypeTriggerText || !committeeTypeOptions) return;

    const currentValue = committeeTypeSelect.value || DEFAULTS['committee-type'];
    const selectedOption = Array.from(committeeTypeSelect.options).find((option) => option.value === currentValue);
    committeeTypeTriggerText.textContent = selectedOption ? selectedOption.textContent : currentValue;
    updateCommitteeTypePreview(currentValue);

    Array.from(committeeTypeOptions.querySelectorAll('.committee-type-option')).forEach((optionEl) => {
      const isSelected = optionEl.dataset.value === currentValue;
      optionEl.setAttribute('aria-selected', String(isSelected));
      optionEl.classList.toggle('is-selected', isSelected);
    });
  }

  function openCommitteeTypeMenu() {
    if (!committeeTypeMenu || !committeeTypeTrigger) return;
    committeeTypeMenu.hidden = false;
    committeeTypeTrigger.setAttribute('aria-expanded', 'true');
    syncCommitteeTypePicker();
  }

  function closeCommitteeTypeMenu() {
    if (!committeeTypeMenu || !committeeTypeTrigger) return;
    committeeTypeMenu.hidden = true;
    committeeTypeTrigger.setAttribute('aria-expanded', 'false');
    syncCommitteeTypePicker();
  }

  function toggleCommitteeTypeMenu() {
    if (!committeeTypeMenu) return;
    if (committeeTypeMenu.hidden) {
      openCommitteeTypeMenu();
    } else {
      closeCommitteeTypeMenu();
    }
  }

  function selectCommitteeType(value) {
    if (!committeeTypeSelect) return;
    committeeTypeSelect.value = value;
    committeeTypeSelect.dispatchEvent(new Event('change', { bubbles: true }));
    syncCommitteeTypePicker();
    closeCommitteeTypeMenu();
    updateFilenamePreview();
    updateCompletion();
  }

  function handleCommitteeTypeKeydown(event) {
    if (!committeeTypeOptions) return;
    const options = Array.from(committeeTypeOptions.querySelectorAll('.committee-type-option'));
    const currentIndex = options.indexOf(document.activeElement);

    if (event.key === 'ArrowDown') {
      event.preventDefault();
      const nextIndex = currentIndex < options.length - 1 ? currentIndex + 1 : 0;
      options[nextIndex]?.focus();
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      const prevIndex = currentIndex > 0 ? currentIndex - 1 : options.length - 1;
      options[prevIndex]?.focus();
    } else if (event.key === 'Home') {
      event.preventDefault();
      options[0]?.focus();
    } else if (event.key === 'End') {
      event.preventDefault();
      options[options.length - 1]?.focus();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      closeCommitteeTypeMenu();
      committeeTypeTrigger?.focus();
    }
  }

  function initCommitteeTypePicker() {
    if (!committeeTypeSelect || !committeeTypeTrigger || !committeeTypeMenu || !committeeTypeOptions) return;

    const optionData = Array.from(committeeTypeSelect.options).map((option) => ({
      value: option.value,
      label: option.textContent || option.value,
      description: getCommitteeTypeDefinition(option.value)
    }));

    committeeTypeOptions.innerHTML = '';

    optionData.forEach((item) => {
      const optionButton = document.createElement('button');
      optionButton.type = 'button';
      optionButton.className = 'committee-type-option';
      optionButton.dataset.value = item.value;
      optionButton.setAttribute('role', 'option');
      optionButton.innerHTML = `
        <span class="committee-type-option__label">${item.label}</span>
        <span class="committee-type-option__meta">Hover for definition</span>
      `;

      optionButton.addEventListener('mouseenter', () => updateCommitteeTypePreview(item.value));
      optionButton.addEventListener('focus', () => updateCommitteeTypePreview(item.value));
      optionButton.addEventListener('click', () => selectCommitteeType(item.value));
      optionButton.addEventListener('keydown', handleCommitteeTypeKeydown);

      committeeTypeOptions.appendChild(optionButton);
    });

    committeeTypeTrigger.addEventListener('click', toggleCommitteeTypeMenu);
    committeeTypeTrigger.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openCommitteeTypeMenu();
        committeeTypeOptions.querySelector('.committee-type-option.is-selected, .committee-type-option')?.focus();
      }
    });

    committeeTypeSelect.addEventListener('change', syncCommitteeTypePicker);

    document.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      const picker = document.getElementById('committee-type-picker');
      if (picker && !picker.contains(target)) {
        closeCommitteeTypeMenu();
      }
    });

    syncCommitteeTypePicker();
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

  function safeFileName(text) {
    const cleaned = String(text || 'Data_Governance_Charter')
      .trim()
      .replace(/[^a-z0-9]+/gi, '_')
      .replace(/^_+|_+$/g, '');
    return cleaned || 'Data_Governance_Charter';
  }

  function setAriaInvalid(id, invalid) {
    const el = getElement(id);
    if (!el) return;
    if (invalid) {
      el.setAttribute('aria-invalid', 'true');
    } else {
      el.removeAttribute('aria-invalid');
    }
  }

  function validateForm() {
    let firstInvalid = null;

    REQUIRED_FIELD_IDS.forEach((id) => {
      const invalid = !getOptionalValue(id);
      setAriaInvalid(id, invalid);
      if (invalid && !firstInvalid) firstInvalid = getElement(id);
    });

    return { isValid: !firstInvalid, firstInvalid };
  }

  function updateFilenamePreview() {
    if (!filenamePreview) return;
    const charterName = getValue('charter-name', DEFAULTS['charter-name']);
    filenamePreview.textContent = `${safeFileName(charterName)}_Charter.docx`;
  }

  function updateCompletion() {
    if (!completionBar || !completionText) return;
    const completed = PROGRESS_FIELD_IDS.filter((id) => getOptionalValue(id)).length;
    const total = PROGRESS_FIELD_IDS.length;
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

  function fillStarterContent() {
    Object.entries(DEFAULTS).forEach(([id, value]) => {
      const field = getElement(id);
      if (!field) return;
      if (!String(field.value || '').trim()) {
        field.value = value;
      }
    });

    syncCommitteeTypePicker();
    updateFilenamePreview();
    updateCompletion();
    setStatus('Starter content loaded. Review and customize before export.', 'success');
  }

  function blankParagraph({ before = 0, after = 120 } = {}) {
    return new docxLib.Paragraph({ text: '', spacing: { before, after } });
  }

  function heading(text, level = docxLib.HeadingLevel.HEADING_1, color) {
    return new docxLib.Paragraph({
      text,
      heading: level,
      spacing: { before: 220, after: 80 },
      children: [new docxLib.TextRun({ text, bold: true, color: color || '1F2933' })]
    });
  }

  function bodyParagraph(text) {
    return new docxLib.Paragraph({
      children: [new docxLib.TextRun({ text })],
      spacing: { after: 120 },
      thematicBreak: false
    });
  }

  function multiParagraphs(text) {
    return toLines(text).map((line) => bodyParagraph(line));
  }

  function bulletList(text) {
    return toLines(text).map(
      (item) =>
        new docxLib.Paragraph({
          text: item,
          bullet: { level: 0 },
          spacing: { after: 80 }
        })
    );
  }

  function tableCell(text, options = {}) {
    return new docxLib.TableCell({
      shading: options.shading ? { fill: options.shading } : undefined,
      verticalAlign: docxLib.VerticalAlign.CENTER,
      children: [
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: String(text || ''),
              bold: Boolean(options.bold),
              color: options.color || '1F2933'
            })
          ],
          spacing: { before: 60, after: 60 }
        })
      ]
    });
  }

  function createKeyValueTable(rows) {
    return new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      layout: docxLib.TableLayoutType.FIXED,
      rows: rows.map(
        ([label, value]) =>
          new docxLib.TableRow({
            children: [
              new docxLib.TableCell({
                width: { size: 28, type: docxLib.WidthType.PERCENTAGE },
                shading: { fill: 'F7F3EC' },
                verticalAlign: docxLib.VerticalAlign.CENTER,
                children: [
                  new docxLib.Paragraph({
                    children: [new docxLib.TextRun({ text: label, bold: true, color: '1F2933' })],
                    spacing: { before: 80, after: 80 }
                  })
                ]
              }),
              new docxLib.TableCell({
                width: { size: 72, type: docxLib.WidthType.PERCENTAGE },
                verticalAlign: docxLib.VerticalAlign.CENTER,
                children: [
                  new docxLib.Paragraph({
                    children: [new docxLib.TextRun({ text: String(value || '') })],
                    spacing: { before: 80, after: 80 }
                  })
                ]
              })
            ]
          })
      ),
      borders: {
        top: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        bottom: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        left: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        right: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        insideHorizontal: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'E8DDD0' },
        insideVertical: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'E8DDD0' }
      }
    });
  }

  function createDataTable(headers, lineText, expectedParts, fallbackText, headerColor, headerTextColor) {
    const lines = toLines(lineText || fallbackText);
    const rows = lines.map((line) => splitWithLimit(line, expectedParts));

    return new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      layout: docxLib.TableLayoutType.FIXED,
      rows: [
        new docxLib.TableRow({
          tableHeader: true,
          children: headers.map((header) =>
            tableCell(header, {
              bold: true,
              shading: headerColor,
              color: headerTextColor
            })
          )
        }),
        ...rows.map(
          (row, rowIndex) =>
            new docxLib.TableRow({
              children: row.map((cellValue) =>
                tableCell(cellValue, { shading: rowIndex % 2 === 0 ? 'FFFFFF' : 'FBF8F2' })
              )
            })
        )
      ],
      borders: {
        top: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        bottom: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        left: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        right: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'C8B9A6' },
        insideHorizontal: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'E8DDD0' },
        insideVertical: { style: docxLib.BorderStyle.SINGLE, size: 1, color: 'E8DDD0' }
      }
    });
  }

  function createSection(title, contentBuilder, color) {
    const content = contentBuilder();
    if (!content || content.length === 0) return [];
    return [heading(title, docxLib.HeadingLevel.HEADING_1, color), ...content, blankParagraph()];
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

    const versionHistoryText = getOptionalValue('version-history') || DEFAULTS['version-history'];

    const versionHistoryTable = createDataTable(
      ['Version', 'Date', 'Author', 'Summary of Changes'],
      versionHistoryText,
      4,
      DEFAULTS['version-history'],
      '5A2D5C',
      'FFFFFF'
    );

    const membersTable = createDataTable(
      ['Name', 'Title', 'Role', 'Voting Status'],
      getValue('members', DEFAULTS.members),
      4,
      DEFAULTS.members,
      '2F5D50',
      'FFFFFF'
    );

    const children = [
      new docxLib.Paragraph({
        children: [new docxLib.TextRun({ text: charterName, bold: true, size: 34, color: '1F2933' })],
        alignment: docxLib.AlignmentType.CENTER,
        spacing: { after: 80 }
      }),
      new docxLib.Paragraph({
        children: [new docxLib.TextRun({ text: agencyName, size: 24, color: '5A6470' })],
        alignment: docxLib.AlignmentType.CENTER,
        spacing: { after: 220 }
      }),
      metadataTable,
      blankParagraph({ after: 60 }),

      ...createSection('1. Purpose', () => multiParagraphs(getValue('purpose', DEFAULTS.purpose)), 'A54A2A'),
      ...createSection('2. Vision / Mission', () => multiParagraphs(getValue('vision-mission', DEFAULTS['vision-mission'])), 'A54A2A'),
      ...createSection('3. Objectives', () => bulletList(getValue('objectives', DEFAULTS.objectives)), 'A54A2A'),
      ...createSection('4. Success Metrics', () => bulletList(getValue('success-metrics', DEFAULTS['success-metrics'])), 'A54A2A'),

      ...createSection(
        '5. Scope & Authority',
        () => [
          heading('In Scope', docxLib.HeadingLevel.HEADING_2, 'A54A2A'),
          ...bulletList(getValue('in-scope', DEFAULTS['in-scope'])),
          heading('Out of Scope', docxLib.HeadingLevel.HEADING_2, 'A54A2A'),
          ...bulletList(getValue('out-of-scope', DEFAULTS['out-of-scope'])),
          heading('Decision Authority', docxLib.HeadingLevel.HEADING_2, 'A54A2A'),
          ...multiParagraphs(getValue('decision-authority', DEFAULTS['decision-authority'])),
          heading('Escalation Path', docxLib.HeadingLevel.HEADING_2, 'A54A2A'),
          ...multiParagraphs(getValue('escalation-path', DEFAULTS['escalation-path']))
        ],
        'A54A2A'
      ),

      ...createSection('6. Guiding Principles', () => bulletList(getValue('guiding-principles', DEFAULTS['guiding-principles'])), '2F5D50'),

      ...createSection(
        '7. Membership & Representation',
        () => [
          heading('Committee Members', docxLib.HeadingLevel.HEADING_2, '2F5D50'),
          membersTable,
          heading('Required Functions / Perspectives', docxLib.HeadingLevel.HEADING_2, '2F5D50'),
          ...bulletList(getValue('required-functions', DEFAULTS['required-functions'])),
          heading('Role Definitions', docxLib.HeadingLevel.HEADING_2, '2F5D50'),
          ...bulletList(getValue('role-definitions', DEFAULTS['role-definitions']))
        ],
        '2F5D50'
      ),

      ...createSection(
        '8. Responsibilities & Deliverables',
        () => [
          heading('Committee Responsibilities', docxLib.HeadingLevel.HEADING_2, '2F5D50'),
          ...bulletList(getValue('responsibilities', DEFAULTS.responsibilities)),
          heading('Annual or Initial Priorities', docxLib.HeadingLevel.HEADING_2, '2F5D50'),
          ...bulletList(getValue('annual-priorities', DEFAULTS['annual-priorities'])),
          heading('Key Deliverables', docxLib.HeadingLevel.HEADING_2, '2F5D50'),
          ...bulletList(getValue('key-deliverables', DEFAULTS['key-deliverables']))
        ],
        '2F5D50'
      ),

      ...createSection(
        '9. Operating Model',
        () => {
          const operatingModelTable = createKeyValueTable([
            ['Meeting Cadence', getValue('meeting-frequency', DEFAULTS['meeting-frequency'])],
            ['Quorum', getValue('quorum', DEFAULTS.quorum)],
            ['Decision-Making Process', getValue('decision-making', DEFAULTS['decision-making'])]
          ]);

          return [
            operatingModelTable,
            heading('Meeting Administration', docxLib.HeadingLevel.HEADING_2, '5A2D5C'),
            ...multiParagraphs(getValue('meeting-administration', DEFAULTS['meeting-administration']))
          ];
        },
        '5A2D5C'
      ),

      ...createSection(
        '10. Policy, Privacy, Security & Sharing',
        () => {
          const policyAlignment = getOptionalValue('policy-alignment');
          const privacySecurity = getOptionalValue('privacy-security-considerations');
          const dataSharing = getOptionalValue('data-sharing');
          const content = [];

          if (policyAlignment) {
            content.push(heading('Policy / Legal / Regulatory Alignment', docxLib.HeadingLevel.HEADING_2, '5A2D5C'));
            content.push(...multiParagraphs(policyAlignment));
          }

          if (privacySecurity) {
            content.push(heading('Privacy, Security & Data Release Considerations', docxLib.HeadingLevel.HEADING_2, '5A2D5C'));
            content.push(...multiParagraphs(privacySecurity));
          }

          if (dataSharing) {
            content.push(heading('Data Sharing & Access Considerations', docxLib.HeadingLevel.HEADING_2, '5A2D5C'));
            content.push(...multiParagraphs(dataSharing));
          }

          return content;
        },
        '5A2D5C'
      ),

      ...createSection('11. Working Groups & Subcommittees', () => {
        const subcommittees = getOptionalValue('subcommittees');
        return subcommittees ? bulletList(subcommittees) : [];
      }, '5A2D5C'),

      ...createSection('12. Version History', () => [versionHistoryTable], '5A2D5C')
    ];

    return new docxLib.Document({
      creator: 'Data Governance Charter Generator',
      title: charterName,
      description: 'Generated charter document for a data governance committee.',
      sections: [
        {
          properties: {},
          children
        }
      ]
    });
  }

  function updateAllHelpers() {
    updateFilenamePreview();
    updateCompletion();
  }

  form.addEventListener('input', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.id && REQUIRED_FIELD_IDS.includes(target.id)) {
      setAriaInvalid(target.id, !getOptionalValue(target.id));
    }

    updateAllHelpers();
  });

  form.addEventListener('reset', () => {
    window.setTimeout(() => {
      REQUIRED_FIELD_IDS.forEach((id) => setAriaInvalid(id, false));
      setStatus('');
      updateAllHelpers();
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

    if (!docxLib || typeof window.saveAs !== 'function') {
      setStatus('Document libraries did not load. Refresh the page and try again.', 'error');
      return;
    }

    try {
      submitButton.disabled = true;
      submitButton.textContent = 'Generating...';
      setStatus('Generating your charter document...', 'success');

      const doc = buildDocument();
      const blob = await docxLib.Packer.toBlob(doc);
      const charterName = getValue('charter-name', DEFAULTS['charter-name']);
      const fileName = `${safeFileName(charterName)}_Charter.docx`;

      window.saveAs(blob, fileName);
      setStatus(`Done. Download started for ${fileName}.`, 'success');
    } catch (error) {
      console.error('Error generating charter:', error);
      setStatus('There was an error generating the charter. Check the browser console for details.', 'error');
    } finally {
      submitButton.disabled = false;
      submitButton.textContent = submitLabel;
    }
  });

  updateAllHelpers();
});
