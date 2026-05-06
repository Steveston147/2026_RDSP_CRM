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
type DeadlineStatus = 'completed' | 'unset' | 'overdue' | 'soon' | 'ok';

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

const isCompleted = (applicant: Applicant) => applicant.passportSubmitted && applicant.enrollmentSubmitted;

const getDateOnly = (date: Date) => new Date(date.getFullYear(), date.getMonth(), date.getDate());

const parseDateInput = (value: string) => {
  if (!value) return null;

  const match = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/) || value.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return null;

  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;

  return date;
};

const getDaysUntilDue = (dueDate: string) => {
  const parsedDueDate = parseDateInput(dueDate);
  if (!parsedDueDate) return null;

  const today = getDateOnly(new Date());
  const due = getDateOnly(parsedDueDate);
  return Math.round((due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
};

const getDeadlineStatus = (applicant: Applicant): DeadlineStatus => {
  if (isCompleted(applicant)) return 'completed';

  const daysUntilDue = getDaysUntilDue(applicant.dueDate);
  if (daysUntilDue === null) return 'unset';
  if (daysUntilDue < 0) return 'overdue';
  if (daysUntilDue <= 7) return 'soon';

  return 'ok';
};

const getDeadlineLabel = (applicant: Applicant) => {
  const status = getDeadlineStatus(applicant);
  const daysUntilDue = getDaysUntilDue(applicant.dueDate);

  if (status === 'completed') return '完了';
  if (status === 'unset') return '期限未設定';
  if (status === 'overdue') return `期限超過 ${Math.abs(daysUntilDue ?? 0)}日`;
  if (status === 'soon') {
    if (daysUntilDue === 0) return '本日期限';
    return `期限間近 あと${daysUntilDue}日`;
  }

  return `期限OK あと${daysUntilDue}日`;
};

const getDeadlineClassName = (applicant: Applicant) => `deadline-badge deadline-${getDeadlineStatus(applicant)}`;

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

  const deadlineStatus = getDeadlineStatus(applicant);
  const deadlineLabel = getDeadlineLabel(applicant);
  const dueDateText = applicant.dueDate ? `提出期限（${applicant.dueDate}）` : '提出期限';
  const requestText = missingItems.join('、');

  if (deadlineStatus === 'overdue') {
    return `${requestText}が未提出で、${deadlineLabel}です。至急リマインドし、必要に応じて個別確認してください。`;
  }

  if (deadlineStatus === 'soon') {
    return `${requestText}が未提出で、${deadlineLabel}です。早めにリマインドし、提出状況を確認してください。`;
  }

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

const createApplicantFromRow = (row: Record<string, unknown>, index: number): Applicant => ({
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
});

const readApplicantsFromExcelFile = async (file: File) => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { raw: true });

  return rows
    .map((row, index) => createApplicantFromRow(row, index))
    .filter((applicant) => applicant.name || applicant.email || applicant.birthDate);
};

const normalizeKeyText = (value: string) => value.trim().toLowerCase();

const getApplicantEmailKey = (applicant: Pick<Applicant, 'email'>) => {
  const email = normalizeKeyText(applicant.email);
  return email ? `email:${email}` : '';
};

const getApplicantNameBirthKey = (applicant: Pick<Applicant, 'name' | 'birthDate'>) => {
  const name = normalizeKeyText(applicant.name);
  const birthDate = normalizeKeyText(applicant.birthDate);

  if (!name || !birthDate) return '';

  return `name:${name}|birth:${birthDate}`;
};

const mergeLatestFormsApplicants = (currentApplicants: Applicant[], latestFormsApplicants: Applicant[]) => {
  const merged = [...currentApplicants];
  const indexByEmail = new Map<string, number>();
  const indexByNameBirth = new Map<string, number>();
  const usedIds = new Set(currentApplicants.map((applicant) => applicant.id));

  currentApplicants.forEach((applicant, index) => {
    const emailKey = getApplicantEmailKey(applicant);
    const nameBirthKey = getApplicantNameBirthKey(applicant);

    if (emailKey) indexByEmail.set(emailKey, index);
    if (nameBirthKey) indexByNameBirth.set(nameBirthKey, index);
  });

  let addedCount = 0;
  let matchedCount = 0;
  let updatedBasicCount = 0;

  latestFormsApplicants.forEach((latestApplicant, latestIndex) => {
    const emailKey = getApplicantEmailKey(latestApplicant);
    const nameBirthKey = getApplicantNameBirthKey(latestApplicant);
    let existingIndex: number | undefined;

    if (emailKey) {
      existingIndex = indexByEmail.get(emailKey);
    }

    if (existingIndex === undefined && nameBirthKey) {
      existingIndex = indexByNameBirth.get(nameBirthKey);
    }

    if (existingIndex !== undefined) {
      const existingApplicant = merged[existingIndex];
      const updatedApplicant = {
        ...existingApplicant,
        name: latestApplicant.name || existingApplicant.name,
        email: latestApplicant.email || existingApplicant.email,
        birthDate: latestApplicant.birthDate || existingApplicant.birthDate,
        nationality: latestApplicant.nationality || existingApplicant.nationality
      };

      const basicInfoChanged =
        updatedApplicant.name !== existingApplicant.name ||
        updatedApplicant.email !== existingApplicant.email ||
        updatedApplicant.birthDate !== existingApplicant.birthDate ||
        updatedApplicant.nationality !== existingApplicant.nationality;

      merged[existingIndex] = updatedApplicant;
      matchedCount += 1;
      if (basicInfoChanged) updatedBasicCount += 1;
      return;
    }

    let newId = latestApplicant.id || `applicant-${currentApplicants.length + latestIndex}`;
    let suffix = 1;
    while (usedIds.has(newId)) {
      newId = `${latestApplicant.id || 'applicant'}-${suffix}`;
      suffix += 1;
    }

    const applicantToAdd = {
      ...latestApplicant,
      id: newId,
      passportSubmitted: latestApplicant.passportSubmitted,
      enrollmentSubmitted: latestApplicant.enrollmentSubmitted,
      dueDate: latestApplicant.dueDate,
      oneDriveLink: latestApplicant.oneDriveLink
    };

    merged.push(applicantToAdd);
    usedIds.add(newId);

    const addedIndex = merged.length - 1;
    if (emailKey) indexByEmail.set(emailKey, addedIndex);
    if (nameBirthKey) indexByNameBirth.set(nameBirthKey, addedIndex);
    addedCount += 1;
  });

  return { merged, addedCount, matchedCount, updatedBasicCount };
};


function App() {
  const [applicants, setApplicants] = useState<Applicant[]>([]);
  const [dueDate, setDueDate] = useState('');
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [selectedApplicantId, setSelectedApplicantId] = useState('');
  const [isApplicantPageOpen, setIsApplicantPageOpen] = useState(false);
  const [reminderApplicantId, setReminderApplicantId] = useState('');
  const [reminderStaffInput, setReminderStaffInput] = useState('');
  const [importMessage, setImportMessage] = useState('');

  const selectedApplicant = useMemo(
    () => applicants.find((applicant) => applicant.id === selectedApplicantId) ?? null,
    [applicants, selectedApplicantId]
  );

  const selectedApplicantIndex = useMemo(
    () => applicants.findIndex((applicant) => applicant.id === selectedApplicantId),
    [applicants, selectedApplicantId]
  );

  const pendingApplicants = useMemo(
    () => applicants.filter((a) => !a.passportSubmitted || !a.enrollmentSubmitted),
    [applicants]
  );

  const urgentApplicants = useMemo(
    () => pendingApplicants.filter((applicant) => ['overdue', 'soon'].includes(getDeadlineStatus(applicant))),
    [pendingApplicants]
  );

  const reminderApplicant = useMemo(() => {
    if (!pendingApplicants.length) return null;

    return pendingApplicants.find((applicant) => applicant.id === reminderApplicantId) ?? pendingApplicants[0];
  }, [pendingApplicants, reminderApplicantId]);

  const pendingCount = pendingApplicants.length;
  const overdueCount = useMemo(
    () => applicants.filter((applicant) => getDeadlineStatus(applicant) === 'overdue').length,
    [applicants]
  );
  const soonCount = useMemo(
    () => applicants.filter((applicant) => getDeadlineStatus(applicant) === 'soon').length,
    [applicants]
  );
  const unsetDueDateCount = useMemo(
    () => applicants.filter((applicant) => getDeadlineStatus(applicant) === 'unset').length,
    [applicants]
  );
  const completedCount = useMemo(() => applicants.filter(isCompleted).length, [applicants]);

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

    const next = await readApplicantsFromExcelFile(file);

    setApplicants(next);
    setSelectedApplicantId(next[0]?.id ?? '');
    setIsApplicantPageOpen(false);
    setReminderApplicantId(next.find((a) => !a.passportSubmitted || !a.enrollmentSubmitted)?.id ?? '');
    setImportMessage(`正本Excelを読み込みました：${next.length}名。画面上の作業内容はこのExcelの内容で置き換わりました。`);
    event.target.value = '';
  };

  const onLatestFormsFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const latestFormsApplicants = await readApplicantsFromExcelFile(file);

    if (!latestFormsApplicants.length) {
      setImportMessage('MS Forms最新Excelを読み込みましたが、追加できる応募者が見つかりませんでした。');
      event.target.value = '';
      return;
    }

    setApplicants((currentApplicants) => {
      const { merged, addedCount, matchedCount, updatedBasicCount } = mergeLatestFormsApplicants(
        currentApplicants,
        latestFormsApplicants
      );
      const nextSelectedApplicantId =
        selectedApplicantId && merged.some((applicant) => applicant.id === selectedApplicantId)
          ? selectedApplicantId
          : merged[0]?.id ?? '';
      const nextReminderApplicantId =
        reminderApplicantId && merged.some((applicant) => applicant.id === reminderApplicantId)
          ? reminderApplicantId
          : merged.find((applicant) => !applicant.passportSubmitted || !applicant.enrollmentSubmitted)?.id ?? '';

      setSelectedApplicantId(nextSelectedApplicantId);
      setReminderApplicantId(nextReminderApplicantId);
      setImportMessage(
        `MS Forms最新Excelを追加反映しました：新規 ${addedCount}名、既存照合 ${matchedCount}名、基本情報更新 ${updatedBasicCount}名。既存の提出状況・期限・OneDriveリンク・メモ・対応履歴・リマインド履歴は維持しています。`
      );

      return merged;
    });

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
      期限状況: getDeadlineLabel(a),
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

  const applyDueDateToUnsetOnly = () => {
    setApplicants((prev) => prev.map((a) => (getDeadlineStatus(a) === 'unset' ? { ...a, dueDate } : a)));
  };

  const applySuggestedNextAction = () => {
    if (!selectedApplicant) return;

    updateApplicant(selectedApplicant.id, {
      nextAction: createNextActionSuggestion(selectedApplicant)
    });
  };

  const openApplicantPage = (applicantId: string) => {
    setSelectedApplicantId(applicantId);
    setIsApplicantPageOpen(true);

    if (typeof window !== 'undefined') {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  const closeApplicantPage = () => {
    setIsApplicantPageOpen(false);
  };

  const openReminderForApplicant = (applicantId: string) => {
    setReminderApplicantId(applicantId);
    setSelectedApplicantId(applicantId);
  };

  const copyReminderEmailForApplicant = async (applicant: Applicant) => {
    const text = createReminderEmail(applicant);

    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    }
  };

  const copyReminderEmail = async () => {
    if (!reminderApplicant) return;
    await copyReminderEmailForApplicant(reminderApplicant);
  };

  const markReminderAsSentForApplicant = (applicant: Applicant) => {
    const today = getTodayInputValue();
    const staff = reminderStaffInput.trim() || applicant.reminderStaff || applicant.staff;
    const nextCount = applicant.reminderCount + 1;
    const missingItems = getMissingDocuments(applicant)
      .map((item) => item.jp)
      .join('、');
    const deadlineInfo = getDeadlineLabel(applicant);
    const reminderEmail = createReminderEmail(applicant);
    const logLine = `${today} ${staff ? `${staff}：` : ''}リマインドメール送信（${nextCount}回目／未提出：${missingItems}／${deadlineInfo}）`;
    const emailLog = [
      '---',
      `Date: ${today}`,
      'Type: Sent reminder email',
      `Staff: ${staff || '未入力'}`,
      'Subject: Reminder for RDSP required documents',
      `Missing documents: ${missingItems || 'なし'}`,
      `Deadline: ${applicant.dueDate || '未設定'} (${deadlineInfo})`,
      '',
      'Body:',
      reminderEmail
    ].join('\n');
    const summaryLine = `${today} ${staff ? `${staff}：` : ''}未提出書類（${missingItems || 'なし'}）のリマインドメールを送信。`;

    updateApplicant(applicant.id, {
      reminderSent: true,
      reminderSentDate: today,
      reminderStaff: staff,
      reminderCount: nextCount,
      reminderNote: appendLine(applicant.reminderNote, logLine),
      staff: staff || applicant.staff,
      responseDate: today,
      responseDetails: appendLine(applicant.responseDetails, logLine),
      pastedEmail: appendLine(applicant.pastedEmail, emailLog),
      emailSummary: appendLine(applicant.emailSummary, summaryLine),
      nextAction: '返信待ち。未提出のまま数日経過した場合は、再リマインドまたは個別確認を行う。'
    });
  };

  const markReminderAsSent = () => {
    if (!reminderApplicant) return;
    markReminderAsSentForApplicant(reminderApplicant);
  };

  const renderApplicantCrmPage = (applicant: Applicant) => {
    const missingDocuments = getMissingDocuments(applicant);
    const applicantNumber = selectedApplicantIndex >= 0 ? selectedApplicantIndex + 1 : '-';
    const hasAnyCommunication =
      applicant.responseDetails.trim() ||
      applicant.reminderNote.trim() ||
      applicant.pastedEmail.trim() ||
      applicant.emailSummary.trim();

    return (
      <section className="crm-page">
        <div className="crm-topbar">
          <button type="button" onClick={closeApplicantPage}>
            ← 応募者一覧へ戻る
          </button>
          <button type="button" onClick={() => openReminderForApplicant(applicant.id)}>
            この学生をリマインド対象にする
          </button>
        </div>

        <section className={`crm-hero deadline-row-${getDeadlineStatus(applicant)}`}>
          <div>
            <p className="crm-kicker">Applicant #{applicantNumber}</p>
            <h2>{applicant.name || '氏名未入力'}</h2>
            <p className="crm-email">{applicant.email || 'メール未入力'}</p>
          </div>
          <div className="crm-hero-badges">
            <span className={getDeadlineClassName(applicant)}>{getDeadlineLabel(applicant)}</span>
            <span className={`crm-completion-pill ${isCompleted(applicant) ? 'is-complete' : 'is-pending'}`}>
              {isCompleted(applicant) ? '書類完了' : '未提出あり'}
            </span>
          </div>
        </section>

        <div className="crm-meta-grid">
          <div>
            <span>生年月日</span>
            <strong>{applicant.birthDate || '未入力'}</strong>
          </div>
          <div>
            <span>国籍</span>
            <strong>{applicant.nationality || '未入力'}</strong>
          </div>
          <div>
            <span>提出期限</span>
            <strong>{applicant.dueDate || '未設定'}</strong>
          </div>
          <div>
            <span>OneDrive</span>
            <strong>{applicant.oneDriveLink.trim() ? '設定済み' : '未設定'}</strong>
          </div>
          <div>
            <span>最終対応日</span>
            <strong>{applicant.responseDate || applicant.reminderSentDate || '未入力'}</strong>
          </div>
          <div>
            <span>担当者</span>
            <strong>{applicant.staff || applicant.reminderStaff || '未入力'}</strong>
          </div>
        </div>

        <div className="crm-layout">
          <div className="crm-main">
            <section className="crm-section-card">
              <h3>提出状況</h3>
              <div className="crm-status-board">
                <label className="crm-check-card">
                  <input
                    type="checkbox"
                    checked={applicant.passportSubmitted}
                    onChange={(e) => updateApplicant(applicant.id, { passportSubmitted: e.target.checked })}
                  />
                  <span>パスポートコピー</span>
                  <strong>{applicant.passportSubmitted ? '提出済み' : '未提出'}</strong>
                </label>
                <label className="crm-check-card">
                  <input
                    type="checkbox"
                    checked={applicant.enrollmentSubmitted}
                    onChange={(e) => updateApplicant(applicant.id, { enrollmentSubmitted: e.target.checked })}
                  />
                  <span>在籍証明書</span>
                  <strong>{applicant.enrollmentSubmitted ? '提出済み' : '未提出'}</strong>
                </label>
                <label>
                  提出期限
                  <input
                    type="date"
                    value={applicant.dueDate}
                    onChange={(e) => updateApplicant(applicant.id, { dueDate: e.target.value })}
                  />
                </label>
                <label>
                  OneDriveリンク
                  <input
                    type="url"
                    placeholder="https://..."
                    value={applicant.oneDriveLink}
                    onChange={(e) => updateApplicant(applicant.id, { oneDriveLink: e.target.value })}
                  />
                </label>
              </div>
            </section>

            <section className="crm-section-card">
              <h3>やり取り履歴・メール記録</h3>
              {hasAnyCommunication ? (
                <div className="crm-timeline">
                  {applicant.responseDetails.trim() && (
                    <article className="crm-timeline-item">
                      <span>対応内容</span>
                      <pre>{applicant.responseDetails}</pre>
                    </article>
                  )}
                  {applicant.reminderNote.trim() && (
                    <article className="crm-timeline-item">
                      <span>リマインド履歴</span>
                      <pre>{applicant.reminderNote}</pre>
                    </article>
                  )}
                  {applicant.emailSummary.trim() && (
                    <article className="crm-timeline-item">
                      <span>メール要約</span>
                      <pre>{applicant.emailSummary}</pre>
                    </article>
                  )}
                  {applicant.pastedEmail.trim() && (
                    <article className="crm-timeline-item">
                      <span>メール本文・貼付記録</span>
                      <pre>{applicant.pastedEmail}</pre>
                    </article>
                  )}
                </div>
              ) : (
                <p className="empty-note">まだメール履歴や対応記録はありません。</p>
              )}
            </section>

            <section className="crm-section-card">
              <h3>記録・修正</h3>
              <div className="crm-form-grid">
                <label>
                  対応内容
                  <textarea
                    rows={4}
                    value={applicant.responseDetails}
                    onChange={(e) => updateApplicant(applicant.id, { responseDetails: e.target.value })}
                  />
                </label>
                <label>
                  リマインドメモ
                  <textarea
                    rows={4}
                    value={applicant.reminderNote}
                    onChange={(e) => updateApplicant(applicant.id, { reminderNote: e.target.value })}
                  />
                </label>
                <label>
                  過去メール・送信メール本文
                  <textarea
                    rows={8}
                    value={applicant.pastedEmail}
                    onChange={(e) => updateApplicant(applicant.id, { pastedEmail: e.target.value })}
                  />
                </label>
                <label>
                  メール要約
                  <textarea
                    rows={5}
                    value={applicant.emailSummary}
                    onChange={(e) => updateApplicant(applicant.id, { emailSummary: e.target.value })}
                  />
                </label>
              </div>
            </section>
          </div>

          <aside className="crm-side">
            <section className="crm-section-card crm-next-action-card">
              <h3>次にやること</h3>
              <p className="crm-suggestion">{createNextActionSuggestion(applicant)}</p>
              <button type="button" onClick={applySuggestedNextAction}>
                提案を反映
              </button>
              <textarea
                rows={5}
                value={applicant.nextAction}
                onChange={(e) => updateApplicant(applicant.id, { nextAction: e.target.value })}
              />
            </section>

            <section className="crm-section-card">
              <h3>メモ・特徴</h3>
              <label>
                メモ
                <textarea
                  rows={5}
                  value={applicant.memo}
                  onChange={(e) => updateApplicant(applicant.id, { memo: e.target.value })}
                />
              </label>
              <label>
                特別リクエスト
                <textarea
                  rows={5}
                  value={applicant.specialRequest}
                  onChange={(e) => updateApplicant(applicant.id, { specialRequest: e.target.value })}
                />
              </label>
            </section>

            <section className="crm-section-card">
              <h3>担当・対応日</h3>
              <label>
                担当者
                <input
                  type="text"
                  value={applicant.staff}
                  onChange={(e) => updateApplicant(applicant.id, { staff: e.target.value })}
                />
              </label>
              <label>
                対応日
                <input
                  type="date"
                  value={applicant.responseDate}
                  onChange={(e) => updateApplicant(applicant.id, { responseDate: e.target.value })}
                />
              </label>
              <label>
                リマインド送信日
                <input
                  type="date"
                  value={applicant.reminderSentDate}
                  onChange={(e) => updateApplicant(applicant.id, { reminderSentDate: e.target.value })}
                />
              </label>
              <label>
                リマインド担当者
                <input
                  type="text"
                  value={applicant.reminderStaff}
                  onChange={(e) => updateApplicant(applicant.id, { reminderStaff: e.target.value })}
                />
              </label>
              <label>
                リマインド回数
                <input
                  type="number"
                  min="0"
                  value={applicant.reminderCount}
                  onChange={(e) => updateApplicant(applicant.id, { reminderCount: Number(e.target.value) || 0 })}
                />
              </label>
            </section>

            {!isCompleted(applicant) && (
              <section className="crm-section-card">
                <h3>リマインド文面</h3>
                <p>
                  未提出：
                  {missingDocuments.length ? missingDocuments.map((item) => item.jp).join('、') : 'なし'}
                </p>
                <textarea readOnly rows={10} value={createReminderEmail(applicant)} className="reminder-textarea" />
                <div className="reminder-actions">
                  <label>
                    送信担当者:
                    <input
                      type="text"
                      placeholder="例：田中"
                      value={reminderStaffInput}
                      onChange={(e) => setReminderStaffInput(e.target.value)}
                    />
                  </label>
                  <button type="button" onClick={() => copyReminderEmailForApplicant(applicant)}>
                    文面をコピー
                  </button>
                  <button type="button" onClick={() => markReminderAsSentForApplicant(applicant)}>
                    リマインド送信済みにする
                  </button>
                </div>
              </section>
            )}
          </aside>
        </div>
      </section>
    );
  };

  return (
    <main className="container">
      <h1>2026 RDSP 応募者進捗管理</h1>
      <p>MS FormsのExcelを読み込み、提出書類の進捗を一元管理します。</p>

      <section className="controls">
        <label>
          正本Excel:
          <input type="file" accept=".xlsx,.xls,.csv" onChange={onFileChange} />
        </label>

        <label>
          MS Forms最新Excelを追加反映:
          <input type="file" accept=".xlsx,.xls,.csv" onChange={onLatestFormsFileChange} />
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
        <button onClick={applyDueDateToUnsetOnly} disabled={!applicants.length || !dueDate || unsetDueDateCount === 0}>
          期限未設定者のみに適用
        </button>
        <button onClick={exportStatus} disabled={!applicants.length}>
          進捗Excelを出力
        </button>
      </section>

      {importMessage && <p className="import-message">{importMessage}</p>}

      {!!applicants.length && (
        <section className="summary-cards">
          <div className="summary-card">
            <span>全体</span>
            <strong>{applicants.length}名</strong>
          </div>
          <div className="summary-card">
            <span>未完了</span>
            <strong>{pendingCount}名</strong>
          </div>
          <div className="summary-card summary-overdue">
            <span>期限超過</span>
            <strong>{overdueCount}名</strong>
          </div>
          <div className="summary-card summary-soon">
            <span>期限間近</span>
            <strong>{soonCount}名</strong>
          </div>
          <div className="summary-card summary-unset">
            <span>期限未設定</span>
            <strong>{unsetDueDateCount}名</strong>
          </div>
          <div className="summary-card summary-completed">
            <span>完了</span>
            <strong>{completedCount}名</strong>
          </div>
        </section>
      )}

      <p>未完了：{pendingCount}名 / 全体：{applicants.length}名</p>
      <p>表示中：{filteredApplicants.length}名</p>

      {isApplicantPageOpen && selectedApplicant ? (
        renderApplicantCrmPage(selectedApplicant)
      ) : (
        <>
          {!!applicants.length && urgentApplicants.length > 0 && (
            <section className="urgent-panel">
              <h2>優先確認リスト</h2>
              <p>期限超過または期限まで7日以内の未完了者です。</p>
              <div className="urgent-list">
                {urgentApplicants.slice(0, 10).map((applicant) => (
                  <button
                    key={applicant.id}
                    type="button"
                    className={`urgent-item deadline-row-${getDeadlineStatus(applicant)}`}
                    onClick={() => openReminderForApplicant(applicant.id)}
                  >
                    <strong>{applicant.name || '氏名未入力'}</strong>
                    <span>{getDeadlineLabel(applicant)}</span>
                    <small>{getMissingDocuments(applicant).map((item) => item.jp).join('、')}</small>
                  </button>
                ))}
              </div>
            </section>
          )}

          {!!applicants.length && (
            <section className="pending-reminders">
              <h2>未提出者リスト・リマインド文面</h2>

              {pendingApplicants.length ? (
                <>
                  <p>未提出者：{pendingApplicants.length}名</p>

                  <div className="reminder-grid">
                    <div className="reminder-list">
                      {pendingApplicants.map((applicant) => {
                        const missingItems = getMissingDocuments(applicant).map((item) => item.jp);

                        return (
                          <div
                            key={applicant.id}
                            className={`reminder-card ${reminderApplicant?.id === applicant.id ? 'is-selected' : ''}`}
                          >
                            <div className="reminder-card-header">
                              <strong>{applicant.name || '氏名未入力'}</strong>
                              <span className={getDeadlineClassName(applicant)}>{getDeadlineLabel(applicant)}</span>
                            </div>
                            <p>未提出：{missingItems.join('、')}</p>
                            <p>
                              期限：{applicant.dueDate || '未設定'} ／ OneDrive：
                              {applicant.oneDriveLink.trim() ? '設定済み' : '未設定'}
                            </p>
                            <p>
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
                          <h3>
                            {reminderApplicant.name} 宛リマインド文面{' '}
                            <span className={getDeadlineClassName(reminderApplicant)}>
                              {getDeadlineLabel(reminderApplicant)}
                            </span>
                          </h3>
                          <textarea
                            readOnly
                            rows={14}
                            value={createReminderEmail(reminderApplicant)}
                            className="reminder-textarea"
                          />

                          <div className="reminder-actions">
                            <label>
                              送信担当者:
                              <input
                                type="text"
                                placeholder="例：田中"
                                value={reminderStaffInput}
                                onChange={(e) => setReminderStaffInput(e.target.value)}
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

                          <p className="reminder-record">
                            記録：{reminderApplicant.reminderCount}回 ／ 最終送信日：
                            {reminderApplicant.reminderSentDate || '未送信'} ／ 担当者：
                            {reminderApplicant.reminderStaff || '未入力'}
                          </p>

                          {reminderApplicant.reminderNote && (
                            <textarea readOnly rows={4} value={reminderApplicant.reminderNote} className="reminder-note" />
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
                <th>No.</th>
                <th>氏名</th>
                <th>メール</th>
                <th>生年月日</th>
                <th>国籍</th>
                <th>パスポートコピー</th>
                <th>在籍証明書</th>
                <th>提出期限</th>
                <th>期限状況</th>
                <th>OneDriveリンク</th>
                <th>リマインド</th>
                <th>詳細</th>
              </tr>
            </thead>
            <tbody>
              {filteredApplicants.map((a) => {
                const applicantNumber = applicants.findIndex((applicant) => applicant.id === a.id) + 1;

                return (
                  <tr key={a.id} className={`deadline-row-${getDeadlineStatus(a)}`}>
                    <td>{applicantNumber || '-'}</td>
                    <td>
                      <button type="button" className="applicant-name-button" onClick={() => openApplicantPage(a.id)}>
                        {a.name || '氏名未入力'}
                      </button>
                    </td>
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
                      <span className={getDeadlineClassName(a)}>{getDeadlineLabel(a)}</span>
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
                      <button type="button" onClick={() => openApplicantPage(a.id)}>
                        個人ページ
                      </button>
                    </td>
                  </tr>
                );
              })}
              {!filteredApplicants.length && (
                <tr>
                  <td colSpan={12}>該当する応募者はいません。</td>
                </tr>
              )}
            </tbody>
          </table>
        </>
      )}
    </main>
  );
}

export default App;
