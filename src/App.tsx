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

const makeId = (row: Record<string, unknown>, index: number) =>
  `${buildApplicantName(row) || getText(row, ['E-mail', 'メール', 'email']) || 'applicant'}-${index}`;

const createNextActionSuggestion = (applicant: Applicant) => {
  const missingItems: string[] = [];

  if (!applicant.passportSubmitted) missingItems.push('パスポートコピー');
  if (!applicant.enrollmentSubmitted) missingItems.push('在籍証明書');

  if (!missingItems.length) {
    if (applicant.specialRequest.trim()) {
      return '書類は完了しています。特別リクエストの内容を確認し、必要に応じて担当者から回答してください。';
    }

    return '書類は完了しています。進捗Excelを出力し、必要に応じて最終確認を行ってください。';
  }

  const dueDateText = applicant.dueDate ? `提出期限（${applicant.dueDate}）` : '提出期限';
  const requestText = missingItems.join('、');

  if (applicant.oneDriveLink.trim()) {
    return `${requestText}の提出依頼メールを作成し、OneDriveリンクと${dueDateText}を案内してください。`;
  }

  return `${requestText}が未提出です。まずOneDriveリンクを設定し、${dueDateText}とあわせて提出依頼メールを送ってください。`;
};

function App() {
  const [applicants, setApplicants] = useState<Applicant[]>([]);
  const [dueDate, setDueDate] = useState('');
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [selectedApplicantId, setSelectedApplicantId] = useState('');

  const selectedApplicant = useMemo(
    () => applicants.find((applicant) => applicant.id === selectedApplicantId) ?? null,
    [applicants, selectedApplicantId]
  );

  const pendingCount = useMemo(
    () => applicants.filter((a) => !a.passportSubmitted || !a.enrollmentSubmitted).length,
    [applicants]
  );

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
      nextAction: getText(row, ['次にやること', '次アクション', 'nextAction'])
    }));

    setApplicants(next);
    setSelectedApplicantId(next[0]?.id ?? '');
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
      次にやること: a.nextAction
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
                <button type="button" onClick={() => setSelectedApplicantId(a.id)}>
                  詳細
                </button>
              </td>
            </tr>
          ))}
          {!filteredApplicants.length && (
            <tr>
              <td colSpan={9}>該当する応募者はいません。</td>
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