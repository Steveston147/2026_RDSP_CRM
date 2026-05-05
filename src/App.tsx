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

const toString = (value: unknown): string => (value == null ? '' : String(value).trim());

const getValue = (row: Record<string, unknown>, keys: string[]): unknown => {
  for (const key of keys) {
    if (key in row && row[key] != null && String(row[key]).trim() !== '') {
      return row[key];
    }
  }
  return '';
};

const formatBirthDate = (value: unknown): string => {
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return '';
    const yyyy = parsed.y.toString().padStart(4, '0');
    const mm = parsed.m.toString().padStart(2, '0');
    const dd = parsed.d.toString().padStart(2, '0');
    return `${yyyy}/${mm}/${dd}`;
  }

  const text = toString(value);
  if (!text) return '';

  const normalized = text.replace(/[-.]/g, '/');
  const matched = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!matched) return normalized;

  const [, y, m, d] = matched;
  return `${y}/${m.padStart(2, '0')}/${d.padStart(2, '0')}`;
};

const toBool = (value: unknown): boolean => {
  const text = toString(value).toLowerCase();
  return text === '提出済み' || text === 'true' || text === '1' || text === 'yes';
};

const makeId = (name: string, index: number) => `${name || 'applicant'}-${index}`;

const createFullName = (row: Record<string, unknown>): string => {
  const first = toString(row['First Name (e.g. John)']);
  const middle = toString(row['Middle Name (e.g. Alan)']);
  const last = toString(row['Last Name (e.g. Smith)']);
  const fullName = [first, middle, last].filter(Boolean).join(' ');
  return fullName || toString(getValue(row, ['名前', '氏名', 'name']));
};

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
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { raw: true });
    if (!rows.length) {
      setApplicants([]);
      return;
    }

    const firstRowKeys = Object.keys(rows[0]);
    const isProgressSheet =
      firstRowKeys.includes('パスポートコピー提出') ||
      firstRowKeys.includes('在籍証明書提出') ||
      firstRowKeys.includes('OneDriveリンク') ||
      firstRowKeys.includes('提出期限') ||
      firstRowKeys.includes('期日');

    const next = rows.map((row, index) => {
      if (isProgressSheet) {
        const name = toString(getValue(row, ['氏名', '名前', 'name']));
        return {
          id: makeId(name, index),
          name,
          email: toString(getValue(row, ['メール', 'E-mail', 'email'])),
          birthDate: formatBirthDate(getValue(row, ['生年月日', 'Date of Birth (yyyy/MM/dd)', 'birthDate'])),
          nationality: toString(
            getValue(row, ['国籍', 'Nationality (Corresponds your passport / e.g. JAPAN)', 'nationality'])
          ),
          passportSubmitted: toBool(getValue(row, ['パスポートコピー提出', 'passportSubmitted'])),
          enrollmentSubmitted: toBool(getValue(row, ['在籍証明書提出', 'enrollmentSubmitted'])),
          dueDate: toString(getValue(row, ['提出期限', '期日', 'dueDate'])),
          oneDriveLink: toString(getValue(row, ['OneDriveリンク', 'oneDriveLink']))
        };
      }

      const name = createFullName(row);
      const answerEmail = toString(getValue(row, ['E-mail']));
      const formEmail = toString(getValue(row, ['メール']));

      return {
        id: makeId(name, index),
        name,
        email: answerEmail || formEmail,
        birthDate: formatBirthDate(getValue(row, ['Date of Birth (yyyy/MM/dd)', '生年月日'])),
        nationality: toString(getValue(row, ['Nationality (Corresponds your passport / e.g. JAPAN)', '国籍'])),
        passportSubmitted: false,
        enrollmentSubmitted: false,
        dueDate: '',
        oneDriveLink: ''
      };
    });

    setApplicants(next);
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
      パスポートコピー提出: a.passportSubmitted ? '提出済み' : '未提出',
      在籍証明書提出: a.enrollmentSubmitted ? '提出済み' : '未提出',
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
