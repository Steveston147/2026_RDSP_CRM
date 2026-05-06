// FILE: src/App.tsx
// PATH: src/App.tsx
import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type Applicant = {
  id: string;
  name: string;
  email: string;
  birthDate: string;
  nationality: string;
  passportSubmitted: boolean;
  enrollmentSubmitted: boolean;
  dueDate: string;
  oneDriveLink: string;
  memo: string;
  specialRequest: string;
  responseDetails: string;
  staff: string;
  responseDate: string;
  pastedEmail: string;
  emailSummary: string;
  nextAction: string;
  reminderSent: boolean;
  reminderSentDate: string;
  reminderStaff: string;
  reminderCount: number;
  reminderNote: string;
};

type StatusFilter = 'all' | 'pending' | 'passportMissing' | 'enrollmentMissing' | 'completed';

const statusFilterOptions: { value: StatusFilter; label: string }[] = [
  { value: 'all', label: '全員' },
  { value: 'pending', label: '未完了のみ' },
  { value: 'passportMissing', label: 'パスポートコピー未提出' },
  { value: 'enrollmentMissing', label: '在籍証明書未提出' },
  { value: 'completed', label: '完了のみ' }
];

const getCell = (row: Record<string, unknown>, keys: string[]) => {
  for (const key of keys) {
    const value = row[key];
    if (value !== undefined && value !== null && String(value).trim() !== '') {
      return value;
    }
  }

  return '';
};

const getText = (row: Record<string, unknown>, keys: string[]) => String(getCell(row, keys)).trim();

const buildApplicantName = (row: Record<string, unknown>) => {
  const firstName = getText(row, ['First Name', 'First Name (e.g. John)', 'firstName']);
  const middleName = getText(row, ['Middle Name', 'Middle Name (e.g. Alan)', 'middleName']);
  const lastName = getText(row, ['Last Name', 'Last Name (e.g. Smith)', 'lastName']);
  const formsName = [firstName, middleName, lastName].filter(Boolean).join(' ');

  return formsName || getText(row, ['氏名', 'name', 'Name']);
};

const formatDateParts = (year: number, month: number, day: number) =>
  [year, month, day].map((part, index) => (index === 0 ? String(part) : String(part).padStart(2, '0'))).join('/');

const getTodayInputValue = () => {
  const today = new Date();
  return [today.getFullYear(), today.getMonth() + 1, today.getDate()]
    .map((part, index) => (index === 0 ? String(part) : String(part).padStart(2, '0')))
    .join('-');
};

const formatBirthDate = (value: unknown) => {
  if (value === undefined || value === null || String(value).trim() === '') return '';

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? formatDateParts(parsed.y, parsed.m, parsed.d) : '';
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDateParts(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  const text = String(value).trim();
  const match = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (match) {
    return formatDateParts(Number(match[1]), Number(match[2]), Number(match[3]));
  }

  return text;
};

const formatInputDate = (value: unknown) => formatBirthDate(value).replace(/\//g, '-');

const parseSubmitted = (value: unknown) => {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value === 1;

  const text = String(value ?? '').trim().toLowerCase();
  return ['提出済み', '済', 'true', 'yes', 'y', '1', 'checked'].includes(text);
};

const parseReminderCount = (value: unknown) => {
  if (typeof value === 'number' && Number.isFinite(value)) return value;

  const text = String(value ?? '').trim();
  if (!text) return 0;

  const numberValue = Number(text.replace(/[^\d]/g, ''));
  return Number.isFinite(numberValue) ? numberValue : 0;
};

const makeId = (row: Record<string, unknown>, index: number) =>
  `${buildApplicantName(row) || getText(row, ['E-mail', 'メール', 'email']) || 'applicant'}-${index}`;

const getMissingDocuments = (applicant: Applicant) => {
  const missing: { jp: string; en: string }[] = [];

  if (!applicant.passportSubmitted) {
    missing.push({ jp: 'パスポートコピー', en: 'Passport copy' });
  }

  if (!applicant.enrollmentSubmitted) {
    missing.push({ jp: '在籍証明書', en: 'Certificate of enrollment' });
  }

  return missing;
};

const getFirstNameForEmail = (name: string) => {
  const first = name.trim().split(/\s+/)[0];
  return first || name || 'there';
};

const formatEnglishDate = (value: string) => {
  if (!value) return '';

  const match = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/) || value.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!match) return value;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  const monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ];

  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return value;

  return `${monthNames[month - 1]} ${day}, ${year}`;
};

const createNextActionSuggestion = (applicant: Applicant) => {
  const missingItems = getMissingDocuments(applicant).map((item) => item.jp);

  if (!missingItems.length) {
    if (applicant.specialRequest.trim()) {
      return '書類は完了しています。特別リクエストの内容を確認し、必要に応じて担当者から回答してください。';
    }

    return '書類は完了しています。進捗Excelを出力し、必要に応じて最終確認を行ってください。';
  }

  const dueDateText = applicant.dueDate ? `提出期限（${applicant.dueDate}）` : '提出期限';
  const requestText = missingItems.join('、');

  if (applicant.reminderSent || applicant.reminderCount > 0) {
    return `${requestText}が未提出です。リマインド送信済みのため、返信を確認し、数日後も未提出なら再リマインドしてください。`;
  }

  if (applicant.oneDriveLink.trim()) {
    return `${requestText}の提出依頼メールを作成し、OneDriveリンクと${dueDateText}を案内してください。`;
  }

  return `${requestText}が未提出です。まずOneDriveリンクを設定し、${dueDateText}とあわせて提出依頼メールを送ってください。`;
};

const createReminderEmail = (applicant: Applicant) => {
  const missingDocuments = getMissingDocuments(applicant);
  const missingList = missingDocuments.map((item) => `- ${item.en}`).join('\n');
  const dueDateLine = applicant.dueDate
    ? `Please upload them by ${formatEnglishDate(applicant.dueDate)}.`
    : 'Please upload them as soon as possible.';
  const oneDriveLine = applicant.oneDriveLink.trim()
    ? `\n\nUpload link:\n${applicant.oneDriveLink.trim()}`
    : '\n\nWe will share the upload link separately if needed.';

  return `Dear ${getFirstNameForEmail(applicant.name)},

This is a gentle reminder that we have not yet received the following document(s):

${missingList}

${dueDateLine}${oneDriveLine}

If you have already submitted them, please kindly ignore this message.

Best regards,
Ritsumeikan Study Abroad Center`;
};

const appendLine = (currentText: string, newLine: string) => {
  const trimmed = currentText.trim();
  return trimmed ? `${trimmed}\n${newLine}` : newLine;
};

function App() {
  const [applicants, setApplicants] = useState<Applicant[]>([]);
  const [dueDate, setDueDate] = useState('');
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [selectedApplicantId, setSelectedApplicantId] = useState('');
  const [reminderApplicantId, setReminderApplicantId] = useState('');
  const [reminderStaffInput, setReminderStaffInput] = useState('');

  const selectedApplicant = useMemo(
    () => applicants.find((applicant) => applicant.id === selectedApplicantId) ?? null,
    [applicants, selectedApplicantId]
  );

  const pendingApplicants = useMemo(
    () => applicants.filter((a) => !a.passportSubmitted || !a.enrollmentSubmitted),
    [applicants]
  );

  const reminderApplicant = useMemo(() => {
    if (!pendingApplicants.length) return null;

    return pendingApplicants.find((applicant) => applicant.id === reminderApplicantId) ?? pendingApplicants[0];
  }, [pendingApplicants, reminderApplicantId]);

  const pendingCount = pendingApplicants.length;

  const filteredApplicants = useMemo(() => {
    const normalizedSearch = searchText.trim().toLowerCase();

    return applicants.filter((a) => {
      const matchesSearch =
        !normalizedSearch ||
        a.name.toLowerCase().includes(normalizedSearch) ||
        a.email.toLowerCase().includes(normalizedSearch);

      if (!matchesSearch) return false;

      switch (statusFilter) {
        case 'pending':
          return !a.passportSubmitted || !a.enrollmentSubmitted;
        case 'passportMissing':
          return !a.passportSubmitted;
        case 'enrollmentMissing':
          return !a.enrollmentSubmitted;
        case 'completed':
          return a.passportSubmitted && a.enrollmentSubmitted;
        default:
          return true;
      }
    });
  }, [applicants, searchText, statusFilter]);

  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { raw: true });

    const next = rows.map((row, index) => ({
      id: makeId(row, index),
      name: buildApplicantName(row),
      email: getText(row, ['E-mail', 'メール', '連絡先メールアドレス', 'email']),
      birthDate: formatBirthDate(getCell(row, ['Date of Birth (yyyy/MM/dd)', 'Date of Birth', '生年月日', 'birthDate'])),
      nationality: getText(row, [
        'Nationality (Corresponds your passport / e.g. JAPAN)',
        'Nationality',
        '国籍',
        'nationality'
      ]),
      passportSubmitted: parseSubmitted(getCell(row, ['パスポートコピー', 'パスポートコピー提出', 'passportSubmitted'])),
      enrollmentSubmitted: parseSubmitted(getCell(row, ['在籍証明書', '在籍証明書提出', 'enrollmentSubmitted'])),
      dueDate: formatInputDate(getCell(row, ['提出期限', '期日', 'dueDate'])),
      oneDriveLink: getText(row, ['OneDriveリンク', 'oneDriveLink']),
      memo: getText(row, ['メモ', 'memo']),
      specialRequest: getText(row, ['特別リクエスト', '特別なリクエスト', 'specialRequest']),
      responseDetails: getText(row, ['対応内容', '対応', 'responseDetails']),
      staff: getText(row, ['担当者', 'staff']),
      responseDate: formatInputDate(getCell(row, ['対応日', 'responseDate'])),
      pastedEmail: getText(row, ['過去メール', '過去メール貼付', 'pastedEmail']),
      emailSummary: getText(row, ['メール要約', 'emailSummary']),
      nextAction: getText(row, ['次にやること', '次アクション', 'nextAction']),
      reminderSent: parseSubmitted(getCell(row, ['リマインド済み', 'reminderSent'])),
      reminderSentDate: formatInputDate(getCell(row, ['リマインド送信日', 'reminderSentDate'])),
      reminderStaff: getText(row, ['リマインド担当者', 'reminderStaff']),
      reminderCount: parseReminderCount(getCell(row, ['リマインド回数', 'reminderCount'])),
      reminderNote: getText(row, ['リマインドメモ', 'reminderNote'])
    }));

    setApplicants(next);
    setSelectedApplicantId(next[0]?.id ?? '');
    setReminderApplicantId(next.find((a) => !a.passportSubmitted || !a.enrollmentSubmitted)?.id ?? '');
    event.target.value = '';
  };

  const updateApplicant = (id: string, patch: Partial<Applicant>) => {
    setApplicants((prev) => prev.map((a) => (a.id === id ? { ...a, ...patch } : a)));
  };

  const exportStatus = () => {
    const rows = applicants.map((a) => ({
      氏名: a.name,
      メール: a.email,
      生年月日: a.birthDate,
      国籍: a.nationality,
      パスポートコピー: a.passportSubmitted ? '提出済み' : '未提出',
      在籍証明書: a.enrollmentSubmitted ? '提出済み' : '未提出',
      提出期限: a.dueDate,
      OneDriveリンク: a.oneDriveLink,
      メモ: a.memo,
      特別リクエスト: a.specialRequest,
      対応内容: a.responseDetails,
      担当者: a.staff,
      対応日: a.responseDate,
      過去メール: a.pastedEmail,
      メール要約: a.emailSummary,
      次にやること: a.nextAction,
      リマインド済み: a.reminderSent ? '済' : '未',
      リマインド送信日: a.reminderSentDate,
      リマインド担当者: a.reminderStaff,
      リマインド回数: a.reminderCount,
      リマインドメモ: a.reminderNote
    }));

    const sheet = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, 'status');
    XLSX.writeFile(wb, 'rdsp_applicant_status.xlsx');
  };

  const applyDueDate = () => {
    setApplicants((prev) => prev.map((a) => ({ ...a, dueDate })));
  };

  const applySuggestedNextAction = () => {
    if (!selectedApplicant) return;

    updateApplicant(selectedApplicant.id, {
      nextAction: createNextActionSuggestion(selectedApplicant)
    });
  };

  const openReminderForApplicant = (applicantId: string) => {
    setReminderApplicantId(applicantId);
    setSelectedApplicantId(applicantId);
  };

  const copyReminderEmail = async () => {
    if (!reminderApplicant) return;

    const text = createReminderEmail(reminderApplicant);

    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    }
  };

  const markReminderAsSent = () => {
    if (!reminderApplicant) return;

    const today = getTodayInputValue();
    const staff = reminderStaffInput.trim() || reminderApplicant.reminderStaff || reminderApplicant.staff;
    const nextCount = reminderApplicant.reminderCount + 1;
    const missingItems = getMissingDocuments(reminderApplicant)
      .map((item) => item.jp)
      .join('、');
    const logLine = `${today} ${staff ? `${staff}：` : ''}リマインドメール送信（${nextCount}回目／未提出：${missingItems}）`;

    updateApplicant(reminderApplicant.id, {
      reminderSent: true,
      reminderSentDate: today,
      reminderStaff: staff,
      reminderCount: nextCount,
      reminderNote: appendLine(reminderApplicant.reminderNote, logLine),
      staff: staff || reminderApplicant.staff,
      responseDate: today,
      responseDetails: appendLine(reminderApplicant.responseDetails, logLine),
      nextAction: '返信待ち。未提出のまま数日経過した場合は、再リマインドまたは個別確認を行う。'
    });
  };

  return (
    <main className="container">
      <h1>2026 RDSP 応募者進捗管理</h1>
      <p>MS FormsのExcelを読み込み、提出書類の進捗を一元管理します。</p>

      <section className="controls">
        <label>
          応募者Excel:
          <input type="file" accept=".xlsx,.xls,.csv" onChange={onFileChange} />
        </label>

        <label>
          検索:
          <input
            type="search"
            placeholder="氏名またはメール"
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
          />
        </label>

        <label>
          ステータス:
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value as StatusFilter)}>
            {statusFilterOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>

        <label>
          共通提出期限:
          <input type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} />
        </label>

        <button onClick={applyDueDate} disabled={!applicants.length || !dueDate}>
          全員に期日を適用
        </button>
        <button onClick={exportStatus} disabled={!applicants.length}>
          進捗Excelを出力
        </button>
      </section>

      <p>未完了：{pendingCount}名 / 全体：{applicants.length}名</p>
      <p>表示中：{filteredApplicants.length}名</p>

      {!!applicants.length && (
        <section
          className="pending-reminders"
          style={{
            margin: '18px 0',
            padding: '16px',
            border: '1px solid #d0d7de',
            borderRadius: '12px',
            background: '#fff'
          }}
        >
          <h2>未提出者リスト・リマインド文面</h2>

          {pendingApplicants.length ? (
            <>
              <p>未提出者：{pendingApplicants.length}名</p>

              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'minmax(260px, 360px) 1fr',
                  gap: '16px',
                  alignItems: 'start'
                }}
              >
                <div
                  style={{
                    maxHeight: '360px',
                    overflow: 'auto',
                    border: '1px solid #e5e7eb',
                    borderRadius: '8px'
                  }}
                >
                  {pendingApplicants.map((applicant) => {
                    const missingItems = getMissingDocuments(applicant).map((item) => item.jp);

                    return (
                      <div
                        key={applicant.id}
                        style={{
                          padding: '10px',
                          borderBottom: '1px solid #e5e7eb',
                          background: reminderApplicant?.id === applicant.id ? '#f0f7ff' : '#fff'
                        }}
                      >
                        <strong>{applicant.name || '氏名未入力'}</strong>
                        <p style={{ margin: '6px 0' }}>未提出：{missingItems.join('、')}</p>
                        <p style={{ margin: '6px 0', fontSize: '13px' }}>
                          期限：{applicant.dueDate || '未設定'} ／ OneDrive：
                          {applicant.oneDriveLink.trim() ? '設定済み' : '未設定'}
                        </p>
                        <p style={{ margin: '6px 0', fontSize: '13px' }}>
                          リマインド：
                          {applicant.reminderCount > 0
                            ? `${applicant.reminderCount}回（最終：${applicant.reminderSentDate || '日付未入力'}）`
                            : '未送信'}
                        </p>
                        <button type="button" onClick={() => openReminderForApplicant(applicant.id)}>
                          文面を表示
                        </button>
                      </div>
                    );
                  })}
                </div>

                <div>
                  {reminderApplicant ? (
                    <>
                      <h3 style={{ marginTop: 0 }}>{reminderApplicant.name} 宛リマインド文面</h3>
                      <textarea
                        readOnly
                        rows={14}
                        value={createReminderEmail(reminderApplicant)}
                        style={{
                          width: '100%',
                          minHeight: '260px',
                          padding: '10px',
                          lineHeight: '1.5'
                        }}
                      />

                      <div
                        style={{
                          display: 'flex',
                          gap: '8px',
                          marginTop: '8px',
                          flexWrap: 'wrap',
                          alignItems: 'center'
                        }}
                      >
                        <label>
                          送信担当者:
                          <input
                            type="text"
                            placeholder="例：田中"
                            value={reminderStaffInput}
                            onChange={(e) => setReminderStaffInput(e.target.value)}
                            style={{ marginLeft: '6px' }}
                          />
                        </label>
                        <button type="button" onClick={copyReminderEmail}>
                          文面をコピー
                        </button>
                        <button type="button" onClick={markReminderAsSent}>
                          リマインド送信済みにする
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            updateApplicant(reminderApplicant.id, {
                              nextAction: createNextActionSuggestion(reminderApplicant)
                            })
                          }
                        >
                          次にやることへ反映
                        </button>
                      </div>

                      <p style={{ margin: '10px 0 0', fontSize: '13px' }}>
                        記録：{reminderApplicant.reminderCount}回 ／ 最終送信日：
                        {reminderApplicant.reminderSentDate || '未送信'} ／ 担当者：
                        {reminderApplicant.reminderStaff || '未入力'}
                      </p>

                      {reminderApplicant.reminderNote && (
                        <textarea
                          readOnly
                          rows={4}
                          value={reminderApplicant.reminderNote}
                          style={{
                            width: '100%',
                            marginTop: '8px',
                            padding: '10px',
                            lineHeight: '1.5'
                          }}
                        />
                      )}
                    </>
                  ) : (
                    <p>左の未提出者から文面を表示する応募者を選んでください。</p>
                  )}
                </div>
              </div>
            </>
          ) : (
            <p>未提出者はいません。全員の書類提出が完了しています。</p>
          )}
        </section>
      )}

      <table>
        <thead>
          <tr>
            <th>氏名</th>
            <th>メール</th>
            <th>生年月日</th>
            <th>国籍</th>
            <th>パスポートコピー</th>
            <th>在籍証明書</th>
            <th>提出期限</th>
            <th>OneDriveリンク</th>
            <th>リマインド</th>
            <th>詳細</th>
          </tr>
        </thead>
        <tbody>
          {filteredApplicants.map((a) => (
            <tr key={a.id}>
              <td>{a.name}</td>
              <td>{a.email}</td>
              <td>{a.birthDate}</td>
              <td>{a.nationality}</td>
              <td>
                <input
                  type="checkbox"
                  checked={a.passportSubmitted}
                  onChange={(e) => updateApplicant(a.id, { passportSubmitted: e.target.checked })}
                />
              </td>
              <td>
                <input
                  type="checkbox"
                  checked={a.enrollmentSubmitted}
                  onChange={(e) => updateApplicant(a.id, { enrollmentSubmitted: e.target.checked })}
                />
              </td>
              <td>
                <input
                  type="date"
                  value={a.dueDate}
                  onChange={(e) => updateApplicant(a.id, { dueDate: e.target.value })}
                />
              </td>
              <td>
                <input
                  type="url"
                  placeholder="https://..."
                  value={a.oneDriveLink}
                  onChange={(e) => updateApplicant(a.id, { oneDriveLink: e.target.value })}
                />
              </td>
              <td>
                {a.reminderCount > 0 ? (
                  <>
                    済 {a.reminderCount}回
                    <br />
                    {a.reminderSentDate}
                  </>
                ) : (
                  '未'
                )}
              </td>
              <td>
                <button type="button" onClick={() => setSelectedApplicantId(a.id)}>
                  詳細
                </button>
              </td>
            </tr>
          ))}
          {!filteredApplicants.length && (
            <tr>
              <td colSpan={10}>該当する応募者はいません。</td>
            </tr>
          )}
        </tbody>
      </table>

      {selectedApplicant && (
        <section
          className="applicant-detail"
          style={{
            marginTop: '24px',
            padding: '16px',
            border: '1px solid #ccc',
            borderRadius: '8px',
            background: '#fafafa'
          }}
        >
          <h2>{selectedApplicant.name} 詳細</h2>

          <p>
            <strong>メール：</strong>
            {selectedApplicant.email || '未入力'}
          </p>
          <p>
            <strong>進捗：</strong>
            パスポートコピー：{selectedApplicant.passportSubmitted ? '提出済み' : '未提出'} ／
            在籍証明書：{selectedApplicant.enrollmentSubmitted ? '提出済み' : '未提出'}
          </p>
          <p>
            <strong>リマインド：</strong>
            {selectedApplicant.reminderCount > 0
              ? `${selectedApplicant.reminderCount}回送信済み（最終：${selectedApplicant.reminderSentDate || '日付未入力'}／担当：${selectedApplicant.reminderStaff || '未入力'}）`
              : '未送信'}
          </p>
          <p>
            <strong>次にやることの提案：</strong>
            {createNextActionSuggestion(selectedApplicant)}
          </p>

          <button type="button" onClick={applySuggestedNextAction}>
            提案を「次にやること」へ反映
          </button>

          <div style={{ display: 'grid', gap: '12px', marginTop: '16px' }}>
            <label>
              メモ
              <textarea
                rows={3}
                value={selectedApplicant.memo}
                onChange={(e) => updateApplicant(selectedApplicant.id, { memo: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              特別リクエスト
              <textarea
                rows={3}
                value={selectedApplicant.specialRequest}
                onChange={(e) => updateApplicant(selectedApplicant.id, { specialRequest: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              対応内容
              <textarea
                rows={3}
                value={selectedApplicant.responseDetails}
                onChange={(e) => updateApplicant(selectedApplicant.id, { responseDetails: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              担当者
              <input
                type="text"
                value={selectedApplicant.staff}
                onChange={(e) => updateApplicant(selectedApplicant.id, { staff: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              対応日
              <input
                type="date"
                value={selectedApplicant.responseDate}
                onChange={(e) => updateApplicant(selectedApplicant.id, { responseDate: e.target.value })}
              />
            </label>

            <label>
              リマインド送信日
              <input
                type="date"
                value={selectedApplicant.reminderSentDate}
                onChange={(e) => updateApplicant(selectedApplicant.id, { reminderSentDate: e.target.value })}
              />
            </label>

            <label>
              リマインド担当者
              <input
                type="text"
                value={selectedApplicant.reminderStaff}
                onChange={(e) => updateApplicant(selectedApplicant.id, { reminderStaff: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              リマインド回数
              <input
                type="number"
                min="0"
                value={selectedApplicant.reminderCount}
                onChange={(e) =>
                  updateApplicant(selectedApplicant.id, { reminderCount: Number(e.target.value) || 0 })
                }
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              リマインドメモ
              <textarea
                rows={4}
                value={selectedApplicant.reminderNote}
                onChange={(e) => updateApplicant(selectedApplicant.id, { reminderNote: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              過去メール貼付
              <textarea
                rows={6}
                value={selectedApplicant.pastedEmail}
                onChange={(e) => updateApplicant(selectedApplicant.id, { pastedEmail: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              メール要約
              <textarea
                rows={4}
                value={selectedApplicant.emailSummary}
                onChange={(e) => updateApplicant(selectedApplicant.id, { emailSummary: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>

            <label>
              次にやること
              <textarea
                rows={4}
                value={selectedApplicant.nextAction}
                onChange={(e) => updateApplicant(selectedApplicant.id, { nextAction: e.target.value })}
                style={{ display: 'block', width: '100%' }}
              />
            </label>
          </div>
        </section>
      )}
    </main>
  );
}

export default App;