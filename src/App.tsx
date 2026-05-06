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
};

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

function App() {
  const [applicants, setApplicants] = useState<Applicant[]>([]);
  const [dueDate, setDueDate] = useState('');

  const pendingCount = useMemo(
    () => applicants.filter((a) => !a.passportSubmitted || !a.enrollmentSubmitted).length,
    [applicants]
  );

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
      oneDriveLink: getText(row, ['OneDriveリンク', 'oneDriveLink'])
    }));

    setApplicants(next);
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
      OneDriveリンク: a.oneDriveLink
    }));

    const sheet = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, 'status');
    XLSX.writeFile(wb, 'rdsp_applicant_status.xlsx');
  };

  const applyDueDate = () => {
    setApplicants((prev) => prev.map((a) => ({ ...a, dueDate })));
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

      <p>未完了: {pendingCount} 名 / 全体: {applicants.length} 名</p>

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
          </tr>
        </thead>
        <tbody>
          {applicants.map((a) => (
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
            </tr>
          ))}
        </tbody>
      </table>
    </main>
  );
}

export default App;
