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
  photoSubmitted: boolean;
  pledgeSubmitted: boolean;
  flightInfoSubmitted: boolean;
  visaDocumentSubmitted: boolean;
  visaRequirementChecked: boolean;
  visaSupportRequired: boolean;
  visaDocumentsSent: boolean;
  visaDocumentsSentDate: string;
  visaDocumentsSentStaff: string;
  paymentConfirmed: boolean;
  insuranceConfirmed: boolean;
  dietaryRestrictionConfirmed: boolean;
  arrivalInfoConfirmed: boolean;
  hasFoodAllergy: boolean;
  allergyEgg: boolean;
  allergyMilk: boolean;
  allergyWheat: boolean;
  allergyBuckwheat: boolean;
  allergyPeanut: boolean;
  allergyTreeNuts: boolean;
  allergyShrimpCrab: boolean;
  allergyFish: boolean;
  allergySoy: boolean;
  allergySesame: boolean;
  allergyOther: string;
  hasReligiousDietaryRestriction: boolean;
  religiousNoPork: boolean;
  religiousNoBeef: boolean;
  religiousNoAlcohol: boolean;
  religiousHalal: boolean;
  religiousVegetarian: boolean;
  religiousVegan: boolean;
  religiousNoGelatin: boolean;
  religiousOther: string;
  dueDate: string;
  oneDriveLink: string;
  memo: string;
  specialRequest: string;
  responseDetails: string;
  staff: string;
  responseDate: string;
  pastedEmail: string;
  emailSummary: string;
  communicationHistory: string;
  nextAction: string;
  reminderSent: boolean;
  reminderSentDate: string;
  reminderStaff: string;
  reminderCount: number;
  reminderNote: string;
};

type StatusFilter =
  | 'all'
  | 'pending'
  | 'passportMissing'
  | 'enrollmentMissing'
  | 'photoMissing'
  | 'pledgeMissing'
  | 'flightInfoMissing'
  | 'visaRequired'
  | 'visaDocumentsNotSent'
  | 'completed';
type DeadlineStatus = 'completed' | 'unset' | 'overdue' | 'soon' | 'ok';
type CommunicationType = 'メール送信' | 'メール受信' | '電話' | '書類確認' | 'メモ' | 'その他';
type PrintMode = 'list' | 'applicant';

type CommunicationDraft = {
  date: string;
  type: CommunicationType;
  staff: string;
  subject: string;
  body: string;
  summary: string;
  nextAction: string;
  reflectNextAction: boolean;
};

type DocumentField =
  | 'passportSubmitted'
  | 'enrollmentSubmitted'
  | 'photoSubmitted'
  | 'pledgeSubmitted'
  | 'flightInfoSubmitted';

type ConfirmationField =
  | 'paymentConfirmed'
  | 'insuranceConfirmed'
  | 'dietaryRestrictionConfirmed'
  | 'arrivalInfoConfirmed';

type AllergyField =
  | 'allergyEgg'
  | 'allergyMilk'
  | 'allergyWheat'
  | 'allergyBuckwheat'
  | 'allergyPeanut'
  | 'allergyTreeNuts'
  | 'allergyShrimpCrab'
  | 'allergyFish'
  | 'allergySoy'
  | 'allergySesame';

type ReligiousDietaryField =
  | 'religiousNoPork'
  | 'religiousNoBeef'
  | 'religiousNoAlcohol'
  | 'religiousHalal'
  | 'religiousVegetarian'
  | 'religiousVegan'
  | 'religiousNoGelatin';

const documentDefinitions: { key: DocumentField; jp: string; en: string }[] = [
  { key: 'passportSubmitted', jp: 'パスポートコピー', en: 'Passport copy' },
  { key: 'enrollmentSubmitted', jp: '在籍証明書', en: 'Certificate of enrollment' },
  { key: 'photoSubmitted', jp: '顔写真', en: 'Photo' },
  { key: 'pledgeSubmitted', jp: '誓約書', en: 'Pledge form' },
  { key: 'flightInfoSubmitted', jp: 'フライト情報', en: 'Flight information' }
];

const confirmationDefinitions: { key: ConfirmationField; label: string }[] = [
  { key: 'paymentConfirmed', label: '支払確認' },
  { key: 'insuranceConfirmed', label: '保険確認' },
  { key: 'dietaryRestrictionConfirmed', label: '食物アレルギー・宗教的食物制限確認' },
  { key: 'arrivalInfoConfirmed', label: '到着情報確認' }
];

const allergyOptions: { key: AllergyField; label: string }[] = [
  { key: 'allergyEgg', label: '卵' },
  { key: 'allergyMilk', label: '乳・乳製品' },
  { key: 'allergyWheat', label: '小麦' },
  { key: 'allergyBuckwheat', label: 'そば' },
  { key: 'allergyPeanut', label: '落花生・ピーナッツ' },
  { key: 'allergyTreeNuts', label: 'ナッツ類' },
  { key: 'allergyShrimpCrab', label: 'えび・かに' },
  { key: 'allergyFish', label: '魚' },
  { key: 'allergySoy', label: '大豆' },
  { key: 'allergySesame', label: 'ごま' }
];

const religiousDietaryOptions: { key: ReligiousDietaryField; label: string }[] = [
  { key: 'religiousNoPork', label: '豚肉を避ける' },
  { key: 'religiousNoBeef', label: '牛肉を避ける' },
  { key: 'religiousNoAlcohol', label: 'アルコールを避ける' },
  { key: 'religiousHalal', label: 'ハラール対応が必要' },
  { key: 'religiousVegetarian', label: 'ベジタリアン' },
  { key: 'religiousVegan', label: 'ヴィーガン' },
  { key: 'religiousNoGelatin', label: 'ゼラチンを避ける' }
];

const communicationTypeOptions: CommunicationType[] = ['メール送信', 'メール受信', '電話', '書類確認', 'メモ', 'その他'];

const statusFilterOptions: { value: StatusFilter; label: string }[] = [
  { value: 'all', label: '全員' },
  { value: 'pending', label: '未完了のみ' },
  { value: 'passportMissing', label: 'パスポートコピー未提出' },
  { value: 'enrollmentMissing', label: '在籍証明書未提出' },
  { value: 'photoMissing', label: '顔写真未提出' },
  { value: 'pledgeMissing', label: '誓約書未提出' },
  { value: 'flightInfoMissing', label: 'フライト情報未提出' },
  { value: 'visaRequired', label: 'ビザ支援書類必要' },
  { value: 'visaDocumentsNotSent', label: 'ビザ支援書類未送付' },
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

const getDateTimeParts = (date = new Date()) => {
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hour = String(date.getHours()).padStart(2, '0');
  const minute = String(date.getMinutes()).padStart(2, '0');

  return { year, month, day, hour, minute };
};

const getExcelExportDateTimeText = (date = new Date()) => {
  const { year, month, day, hour, minute } = getDateTimeParts(date);
  return `${year}/${month}/${day} ${hour}:${minute}`;
};

const getExcelExportTimestamp = (date = new Date()) => {
  const { year, month, day, hour, minute } = getDateTimeParts(date);
  return `${year}${month}${day}_${hour}${minute}`;
};

const sanitizeFileNamePart = (value: string) =>
  value
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '')
    .replace(/\s+/g, '_')
    .slice(0, 40);

const createExportFileName = (operatorName: string, date = new Date()) => {
  const timestamp = getExcelExportTimestamp(date);
  const safeOperatorName = sanitizeFileNamePart(operatorName);

  return `RDSP_Applicant_Status_${timestamp}${safeOperatorName ? `_${safeOperatorName}` : ''}.xlsx`;
};

const createInitialCommunicationDraft = (): CommunicationDraft => ({
  date: getTodayInputValue(),
  type: 'メール受信',
  staff: '',
  subject: '',
  body: '',
  summary: '',
  nextAction: '',
  reflectNextAction: false
});

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

const parseAffirmative = (value: unknown) => {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value === 1;

  const text = String(value ?? '').trim().toLowerCase();
  return ['あり', '有', '有り', '該当', '提出済み', '済', 'true', 'yes', 'y', '1', 'checked'].includes(text);
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

const getMissingDocuments = (applicant: Applicant) =>
  documentDefinitions
    .filter((definition) => !applicant[definition.key])
    .map((definition) => ({ jp: definition.jp, en: definition.en }));

const getMissingConfirmations = (applicant: Applicant) =>
  confirmationDefinitions.filter((definition) => !applicant[definition.key]).map((definition) => definition.label);

const isCompleted = (applicant: Applicant) => documentDefinitions.every((definition) => applicant[definition.key]);

const getDocumentCompletionLabel = (applicant: Applicant) => {
  const missingCount = getMissingDocuments(applicant).length;
  return missingCount ? `${missingCount}件未完了` : '完了';
};

const getConfirmationCompletionLabel = (applicant: Applicant) => {
  const missingCount = getMissingConfirmations(applicant).length;
  return missingCount ? `${missingCount}件未確認` : '確認済み';
};

const getSelectedFoodAllergies = (applicant: Applicant) => {
  if (!applicant.hasFoodAllergy) return [];

  const selected = allergyOptions.filter((option) => applicant[option.key]).map((option) => option.label);
  const other = applicant.allergyOther.trim();
  return other ? [...selected, `その他：${other}`] : selected;
};

const getSelectedReligiousDietaryRestrictions = (applicant: Applicant) => {
  if (!applicant.hasReligiousDietaryRestriction) return [];

  const selected = religiousDietaryOptions.filter((option) => applicant[option.key]).map((option) => option.label);
  const other = applicant.religiousOther.trim();
  return other ? [...selected, `その他：${other}`] : selected;
};

const getFoodAllergySummary = (applicant: Applicant) => {
  if (!applicant.hasFoodAllergy) return applicant.dietaryRestrictionConfirmed ? 'なし' : '未確認';

  const selected = getSelectedFoodAllergies(applicant);
  return selected.length ? selected.join('、') : 'あり（詳細未入力）';
};

const getReligiousDietarySummary = (applicant: Applicant) => {
  if (!applicant.hasReligiousDietaryRestriction) return applicant.dietaryRestrictionConfirmed ? 'なし' : '未確認';

  const selected = getSelectedReligiousDietaryRestrictions(applicant);
  return selected.length ? selected.join('、') : 'あり（詳細未入力）';
};

const getVisaStatusLabel = (applicant: Applicant) => {
  if (!applicant.visaRequirementChecked) return '要否未確認';
  if (!applicant.visaSupportRequired) return '不要';
  if (applicant.visaDocumentsSent) {
    return `必要・送付済み${applicant.visaDocumentsSentDate ? `（${applicant.visaDocumentsSentDate}）` : ''}`;
  }

  return '必要・未送付';
};

const getVisaStatusClassName = (applicant: Applicant) => {
  if (!applicant.visaRequirementChecked) return 'visa-status-pill visa-unchecked';
  if (!applicant.visaSupportRequired) return 'visa-status-pill visa-not-required';
  if (applicant.visaDocumentsSent) return 'visa-status-pill visa-sent';
  return 'visa-status-pill visa-pending';
};

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
    if (!applicant.visaRequirementChecked) {
      return '書類は完了しています。国籍を確認し、ビザ支援書類が必要かどうかを職員側で確認してください。';
    }

    if (applicant.visaSupportRequired && !applicant.visaDocumentsSent) {
      return '書類は完了しています。ビザ支援書類が必要です。必要書類を作成・送付し、送付済みにチェックしてください。';
    }

    const missingConfirmations = getMissingConfirmations(applicant);

    if (missingConfirmations.length) {
      return `書類は完了しています。次に、${missingConfirmations.join('、')}を確認してください。`;
    }

    if (applicant.hasFoodAllergy || applicant.hasReligiousDietaryRestriction) {
      return `書類は完了しています。食事制限（アレルギー：${getFoodAllergySummary(applicant)}／宗教的制限：${getReligiousDietarySummary(applicant)}）を関係者へ共有してください。`;
    }

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

const createCommunicationRecordText = (draft: CommunicationDraft, applicant: Applicant) => {
  const date = draft.date || getTodayInputValue();
  const staff = draft.staff.trim() || applicant.staff || applicant.reminderStaff || '未入力';
  const subject = draft.subject.trim();
  const summary = draft.summary.trim();
  const body = draft.body.trim();
  const nextAction = draft.nextAction.trim();

  return [
    '---',
    `Date: ${date}`,
    `Type: ${draft.type}`,
    `Staff: ${staff}`,
    subject ? `Subject: ${subject}` : '',
    summary ? `Summary: ${summary}` : '',
    nextAction ? `Next Action: ${nextAction}` : '',
    body ? '' : '',
    body ? 'Body / Note:' : '',
    body
  ]
    .filter((line, index, lines) => line !== '' || (index > 0 && index < lines.length - 1))
    .join('\n')
    .trim();
};

const createCommunicationSummaryLine = (draft: CommunicationDraft, applicant: Applicant) => {
  const date = draft.date || getTodayInputValue();
  const staff = draft.staff.trim() || applicant.staff || applicant.reminderStaff || '';
  const subject = draft.subject.trim();
  const summary = draft.summary.trim() || draft.body.trim().slice(0, 80);

  return `${date} ${staff ? `${staff}：` : ''}${draft.type}${subject ? `（${subject}）` : ''}${summary ? `：${summary}` : ''}`;
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
  photoSubmitted: parseSubmitted(getCell(row, ['顔写真', '顔写真提出', 'photoSubmitted'])),
  pledgeSubmitted: parseSubmitted(getCell(row, ['誓約書', '誓約書提出', 'pledgeSubmitted'])),
  flightInfoSubmitted: parseSubmitted(getCell(row, ['フライト情報', 'フライト情報提出', 'flightInfoSubmitted'])),
  visaDocumentSubmitted: parseSubmitted(getCell(row, ['ビザ書類', 'ビザ書類提出', 'visaDocumentSubmitted'])),
  visaRequirementChecked:
    parseSubmitted(getCell(row, ['ビザ要否確認', 'ビザ要否確認済み', 'visaRequirementChecked'])) ||
    parseAffirmative(getCell(row, ['ビザ支援書類必要', 'ビザ必要', 'visaSupportRequired'])) ||
    parseSubmitted(getCell(row, ['ビザ支援書類送付済み', 'ビザ書類送付済み', 'visaDocumentsSent'])) ||
    parseSubmitted(getCell(row, ['ビザ書類', 'ビザ書類提出', 'visaDocumentSubmitted'])),
  visaSupportRequired:
    parseAffirmative(getCell(row, ['ビザ支援書類必要', 'ビザ必要', 'visaSupportRequired'])) ||
    parseSubmitted(getCell(row, ['ビザ書類', 'ビザ書類提出', 'visaDocumentSubmitted'])),
  visaDocumentsSent:
    parseSubmitted(getCell(row, ['ビザ支援書類送付済み', 'ビザ書類送付済み', 'visaDocumentsSent'])) ||
    parseSubmitted(getCell(row, ['ビザ書類', 'ビザ書類提出', 'visaDocumentSubmitted'])),
  visaDocumentsSentDate: formatInputDate(getCell(row, ['ビザ支援書類送付日', 'ビザ書類送付日', 'visaDocumentsSentDate'])),
  visaDocumentsSentStaff: getText(row, ['ビザ支援書類送付担当者', 'ビザ書類送付担当者', 'visaDocumentsSentStaff']),
  paymentConfirmed: parseSubmitted(getCell(row, ['支払確認', 'paymentConfirmed'])),
  insuranceConfirmed: parseSubmitted(getCell(row, ['保険確認', 'insuranceConfirmed'])),
  dietaryRestrictionConfirmed: parseSubmitted(getCell(row, ['食事制限確認', '食物アレルギー・宗教的食物制限確認', 'dietaryRestrictionConfirmed'])),
  arrivalInfoConfirmed: parseSubmitted(getCell(row, ['到着情報確認', 'arrivalInfoConfirmed'])),
  hasFoodAllergy: parseAffirmative(getCell(row, ['食物アレルギー有無', '食物アレルギーあり', 'hasFoodAllergy'])),
  allergyEgg: parseAffirmative(getCell(row, ['アレルギー_卵', '卵アレルギー', 'allergyEgg'])),
  allergyMilk: parseAffirmative(getCell(row, ['アレルギー_乳', '乳アレルギー', 'allergyMilk'])),
  allergyWheat: parseAffirmative(getCell(row, ['アレルギー_小麦', '小麦アレルギー', 'allergyWheat'])),
  allergyBuckwheat: parseAffirmative(getCell(row, ['アレルギー_そば', 'そばアレルギー', 'allergyBuckwheat'])),
  allergyPeanut: parseAffirmative(getCell(row, ['アレルギー_落花生', 'ピーナッツアレルギー', 'allergyPeanut'])),
  allergyTreeNuts: parseAffirmative(getCell(row, ['アレルギー_ナッツ類', 'ナッツ類アレルギー', 'allergyTreeNuts'])),
  allergyShrimpCrab: parseAffirmative(getCell(row, ['アレルギー_えびかに', 'えび・かにアレルギー', 'allergyShrimpCrab'])),
  allergyFish: parseAffirmative(getCell(row, ['アレルギー_魚', '魚アレルギー', 'allergyFish'])),
  allergySoy: parseAffirmative(getCell(row, ['アレルギー_大豆', '大豆アレルギー', 'allergySoy'])),
  allergySesame: parseAffirmative(getCell(row, ['アレルギー_ごま', 'ごまアレルギー', 'allergySesame'])),
  allergyOther: getText(row, ['アレルギー_その他', '食物アレルギーその他', 'allergyOther']),
  hasReligiousDietaryRestriction: parseAffirmative(getCell(row, ['宗教的食物制限有無', '宗教的食物制限あり', 'hasReligiousDietaryRestriction'])),
  religiousNoPork: parseAffirmative(getCell(row, ['宗教食_豚肉不可', '豚肉を避ける', 'religiousNoPork'])),
  religiousNoBeef: parseAffirmative(getCell(row, ['宗教食_牛肉不可', '牛肉を避ける', 'religiousNoBeef'])),
  religiousNoAlcohol: parseAffirmative(getCell(row, ['宗教食_アルコール不可', 'アルコールを避ける', 'religiousNoAlcohol'])),
  religiousHalal: parseAffirmative(getCell(row, ['宗教食_ハラール', 'ハラール対応', 'religiousHalal'])),
  religiousVegetarian: parseAffirmative(getCell(row, ['宗教食_ベジタリアン', 'ベジタリアン', 'religiousVegetarian'])),
  religiousVegan: parseAffirmative(getCell(row, ['宗教食_ヴィーガン', 'ヴィーガン', 'religiousVegan'])),
  religiousNoGelatin: parseAffirmative(getCell(row, ['宗教食_ゼラチン不可', 'ゼラチンを避ける', 'religiousNoGelatin'])),
  religiousOther: getText(row, ['宗教食_その他', '宗教的食物制限その他', 'religiousOther']),
  dueDate: formatInputDate(getCell(row, ['提出期限', '期日', 'dueDate'])),
  oneDriveLink: getText(row, ['OneDriveリンク', 'oneDriveLink']),
  memo: getText(row, ['メモ', 'memo']),
  specialRequest: getText(row, ['特別リクエスト', '特別なリクエスト', 'specialRequest']),
  responseDetails: getText(row, ['対応内容', '対応', 'responseDetails']),
  staff: getText(row, ['担当者', 'staff']),
  responseDate: formatInputDate(getCell(row, ['対応日', 'responseDate'])),
  pastedEmail: getText(row, ['過去メール', '過去メール貼付', 'pastedEmail']),
  emailSummary: getText(row, ['メール要約', 'emailSummary']),
  communicationHistory: getText(row, ['やり取り履歴', 'Communication History', 'communicationHistory']),
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
  const [operatorName, setOperatorName] = useState('');
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [selectedApplicantId, setSelectedApplicantId] = useState('');
  const [isApplicantPageOpen, setIsApplicantPageOpen] = useState(false);
  const [reminderApplicantId, setReminderApplicantId] = useState('');
  const [reminderStaffInput, setReminderStaffInput] = useState('');
  const [communicationDraft, setCommunicationDraft] = useState<CommunicationDraft>(() => createInitialCommunicationDraft());
  const [importMessage, setImportMessage] = useState('');

  const selectedApplicant = useMemo(
    () => applicants.find((applicant) => applicant.id === selectedApplicantId) ?? null,
    [applicants, selectedApplicantId]
  );

  const selectedApplicantIndex = useMemo(
    () => applicants.findIndex((applicant) => applicant.id === selectedApplicantId),
    [applicants, selectedApplicantId]
  );

  const exportFileNamePreview = useMemo(() => createExportFileName(operatorName), [operatorName]);

  const pendingApplicants = useMemo(() => applicants.filter((a) => !isCompleted(a)), [applicants]);

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
  const confirmationPendingCount = useMemo(
    () => applicants.filter((applicant) => getMissingConfirmations(applicant).length > 0).length,
    [applicants]
  );
  const visaRequirementUncheckedCount = useMemo(
    () => applicants.filter((applicant) => !applicant.visaRequirementChecked).length,
    [applicants]
  );
  const visaDocumentsPendingCount = useMemo(
    () => applicants.filter((applicant) => applicant.visaSupportRequired && !applicant.visaDocumentsSent).length,
    [applicants]
  );

  const dietaryAttentionApplicants = useMemo(
    () => applicants.filter((applicant) => applicant.hasFoodAllergy || applicant.hasReligiousDietaryRestriction),
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
          return !isCompleted(a);
        case 'passportMissing':
          return !a.passportSubmitted;
        case 'enrollmentMissing':
          return !a.enrollmentSubmitted;
        case 'photoMissing':
          return !a.photoSubmitted;
        case 'pledgeMissing':
          return !a.pledgeSubmitted;
        case 'flightInfoMissing':
          return !a.flightInfoSubmitted;
        case 'visaRequired':
          return a.visaSupportRequired;
        case 'visaDocumentsNotSent':
          return a.visaSupportRequired && !a.visaDocumentsSent;
        case 'completed':
          return isCompleted(a);
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
    setReminderApplicantId(next.find((a) => !isCompleted(a))?.id ?? '');
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
          : merged.find((applicant) => !isCompleted(applicant))?.id ?? '';

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
    const exportedAt = new Date();
    const exportDateTimeText = getExcelExportDateTimeText(exportedAt);
    const exportOperatorName = operatorName.trim();
    const fileName = createExportFileName(exportOperatorName, exportedAt);

    const rows = applicants.map((a) => ({
      氏名: a.name,
      メール: a.email,
      生年月日: a.birthDate,
      国籍: a.nationality,
      パスポートコピー: a.passportSubmitted ? '提出済み' : '未提出',
      在籍証明書: a.enrollmentSubmitted ? '提出済み' : '未提出',
      顔写真: a.photoSubmitted ? '提出済み' : '未提出',
      誓約書: a.pledgeSubmitted ? '提出済み' : '未提出',
      フライト情報: a.flightInfoSubmitted ? '提出済み' : '未提出',
      ビザ要否確認: a.visaRequirementChecked ? '済' : '未',
      ビザ支援書類必要: a.visaSupportRequired ? '必要' : a.visaRequirementChecked ? '不要' : '未確認',
      ビザ支援書類送付済み: a.visaDocumentsSent ? '済' : '未',
      ビザ支援書類送付日: a.visaDocumentsSentDate,
      ビザ支援書類送付担当者: a.visaDocumentsSentStaff,
      ビザ状況: getVisaStatusLabel(a),
      書類完了状況: getDocumentCompletionLabel(a),
      支払確認: a.paymentConfirmed ? '済' : '未',
      保険確認: a.insuranceConfirmed ? '済' : '未',
      食物アレルギー宗教的食物制限確認: a.dietaryRestrictionConfirmed ? '済' : '未',
      到着情報確認: a.arrivalInfoConfirmed ? '済' : '未',
      確認事項状況: getConfirmationCompletionLabel(a),
      食物アレルギー有無: a.hasFoodAllergy ? 'あり' : 'なし',
      食物アレルギー内容: getFoodAllergySummary(a),
      宗教的食物制限有無: a.hasReligiousDietaryRestriction ? 'あり' : 'なし',
      宗教的食物制限内容: getReligiousDietarySummary(a),
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
      やり取り履歴: a.communicationHistory,
      次にやること: a.nextAction,
      リマインド済み: a.reminderSent ? '済' : '未',
      リマインド送信日: a.reminderSentDate,
      リマインド担当者: a.reminderStaff,
      リマインド回数: a.reminderCount,
      リマインドメモ: a.reminderNote,
      出力作業者: exportOperatorName,
      出力日時: exportDateTimeText
    }));

    const exportInfoRows = [
      { 項目: '出力日時', 内容: exportDateTimeText },
      { 項目: '出力作業者', 内容: exportOperatorName || '未入力' },
      { 項目: '応募者数', 内容: `${applicants.length}名` },
      { 項目: '未完了', 内容: `${pendingCount}名` },
      { 項目: '確認事項未完了', 内容: `${confirmationPendingCount}名` },
      { 項目: 'ビザ要否未確認', 内容: `${visaRequirementUncheckedCount}名` },
      { 項目: 'ビザ支援書類未送付', 内容: `${visaDocumentsPendingCount}名` },
      { 項目: '期限超過', 内容: `${overdueCount}名` },
      { 項目: '期限間近', 内容: `${soonCount}名` },
      { 項目: '期限未設定', 内容: `${unsetDueDateCount}名` },
      { 項目: '完了', 内容: `${completedCount}名` }
    ];

    const sheet = XLSX.utils.json_to_sheet(rows);
    const exportInfoSheet = XLSX.utils.json_to_sheet(exportInfoRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet, 'status');
    XLSX.utils.book_append_sheet(wb, exportInfoSheet, 'export_info');
    XLSX.writeFile(wb, fileName);

    setImportMessage(
      `進捗Excelを出力しました：${fileName}（作業者：${exportOperatorName || '未入力'}／出力日時：${exportDateTimeText}）`
    );
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
    const applicant = applicants.find((a) => a.id === applicantId);

    setSelectedApplicantId(applicantId);
    setCommunicationDraft({
      ...createInitialCommunicationDraft(),
      staff: applicant?.staff || applicant?.reminderStaff || reminderStaffInput || operatorName
    });
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

  const updateCommunicationDraft = (patch: Partial<CommunicationDraft>) => {
    setCommunicationDraft((prev) => ({ ...prev, ...patch }));
  };

  const addCommunicationRecord = (applicant: Applicant) => {
    const hasContent =
      communicationDraft.subject.trim() ||
      communicationDraft.body.trim() ||
      communicationDraft.summary.trim() ||
      communicationDraft.nextAction.trim();

    if (!hasContent) return;

    const recordText = createCommunicationRecordText(communicationDraft, applicant);
    const summaryLine = createCommunicationSummaryLine(communicationDraft, applicant);
    const staff = communicationDraft.staff.trim() || operatorName.trim() || applicant.staff || applicant.reminderStaff;
    const date = communicationDraft.date || getTodayInputValue();
    const isEmailRecord = communicationDraft.type === 'メール送信' || communicationDraft.type === 'メール受信';
    const emailRecord = isEmailRecord ? recordText : '';

    updateApplicant(applicant.id, {
      communicationHistory: appendLine(applicant.communicationHistory, recordText),
      responseDetails: appendLine(applicant.responseDetails, summaryLine),
      emailSummary: isEmailRecord ? appendLine(applicant.emailSummary, summaryLine) : applicant.emailSummary,
      pastedEmail: emailRecord ? appendLine(applicant.pastedEmail, emailRecord) : applicant.pastedEmail,
      staff: staff || applicant.staff,
      responseDate: date,
      nextAction: communicationDraft.reflectNextAction && communicationDraft.nextAction.trim() ? communicationDraft.nextAction.trim() : applicant.nextAction
    });

    setCommunicationDraft({
      ...createInitialCommunicationDraft(),
      staff
    });
  };

  const markReminderAsSentForApplicant = (applicant: Applicant) => {
    const today = getTodayInputValue();
    const staff = reminderStaffInput.trim() || operatorName.trim() || applicant.reminderStaff || applicant.staff;
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
      communicationHistory: appendLine(applicant.communicationHistory, emailLog),
      pastedEmail: appendLine(applicant.pastedEmail, emailLog),
      emailSummary: appendLine(applicant.emailSummary, summaryLine),
      nextAction: '返信待ち。未提出のまま数日経過した場合は、再リマインドまたは個別確認を行う。'
    });
  };

  const markReminderAsSent = () => {
    if (!reminderApplicant) return;
    markReminderAsSentForApplicant(reminderApplicant);
  };

  const markVisaDocumentsAsSentForApplicant = (applicant: Applicant) => {
    const today = getTodayInputValue();
    const staff = operatorName.trim() || applicant.visaDocumentsSentStaff || applicant.staff || applicant.reminderStaff;
    const logLine = `${today} ${staff ? `${staff}：` : ''}ビザ支援書類を送付済みとして記録。`;

    updateApplicant(applicant.id, {
      visaRequirementChecked: true,
      visaSupportRequired: true,
      visaDocumentsSent: true,
      visaDocumentsSentDate: today,
      visaDocumentsSentStaff: staff,
      staff: staff || applicant.staff,
      responseDate: today,
      responseDetails: appendLine(applicant.responseDetails, logLine),
      communicationHistory: appendLine(applicant.communicationHistory, logLine),
      nextAction: 'ビザ支援書類送付済み。必要に応じて、学生からの受領確認や追加質問を確認してください。'
    });
  };

  const printWithMode = (mode: PrintMode) => {
    if (typeof window === 'undefined' || typeof document === 'undefined') return;

    document.body.setAttribute('data-print-mode', mode);
    const cleanup = () => document.body.removeAttribute('data-print-mode');
    window.addEventListener('afterprint', cleanup, { once: true });
    window.setTimeout(() => {
      window.print();
      window.setTimeout(cleanup, 1000);
    }, 80);
  };

  const printApplicantList = () => {
    setIsApplicantPageOpen(false);
    window.setTimeout(() => printWithMode('list'), 120);
  };

  const printCurrentApplicantPage = () => {
    if (!selectedApplicant) return;
    setIsApplicantPageOpen(true);
    window.setTimeout(() => printWithMode('applicant'), 120);
  };

  const renderApplicantCrmPage = (applicant: Applicant) => {
    const missingDocuments = getMissingDocuments(applicant);
    const applicantNumber = selectedApplicantIndex >= 0 ? selectedApplicantIndex + 1 : '-';
    const hasAnyCommunication =
      applicant.communicationHistory.trim() ||
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
          <button type="button" onClick={printCurrentApplicantPage}>
            この個人ページを印刷
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
            <span>ビザ状況</span>
            <strong>{getVisaStatusLabel(applicant)}</strong>
          </div>
          <div>
            <span>食物アレルギー</span>
            <strong>{getFoodAllergySummary(applicant)}</strong>
          </div>
          <div>
            <span>宗教的食物制限</span>
            <strong>{getReligiousDietarySummary(applicant)}</strong>
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
            <section className="crm-section-card print-section-documents">
              <h3>提出状況</h3>
              <div className="crm-status-board">
                {documentDefinitions.map((definition) => (
                  <label key={definition.key} className="crm-check-card">
                    <input
                      type="checkbox"
                      checked={applicant[definition.key]}
                      onChange={(e) =>
                        updateApplicant(applicant.id, { [definition.key]: e.target.checked } as Partial<Applicant>)
                      }
                    />
                    <span>{definition.jp}</span>
                    <strong>{applicant[definition.key] ? '提出済み' : '未提出'}</strong>
                  </label>
                ))}
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

            <section className="crm-section-card visa-section-card print-section-visa">
              <h3>ビザ確認・送付管理</h3>
              <p className="empty-note">国籍によりビザ支援書類が必要かどうかを職員側で確認し、送付したら送付済みとして記録します。</p>
              <div className="crm-status-board">
                <label className="crm-check-card">
                  <input
                    type="checkbox"
                    checked={applicant.visaRequirementChecked}
                    onChange={(e) =>
                      updateApplicant(applicant.id, {
                        visaRequirementChecked: e.target.checked,
                        visaSupportRequired: e.target.checked ? applicant.visaSupportRequired : false,
                        visaDocumentsSent: e.target.checked ? applicant.visaDocumentsSent : false
                      })
                    }
                  />
                  <span>ビザ要否を確認済み</span>
                  <strong>{applicant.visaRequirementChecked ? '確認済み' : '未確認'}</strong>
                </label>
                <label className="crm-check-card">
                  <input
                    type="checkbox"
                    checked={applicant.visaSupportRequired}
                    onChange={(e) =>
                      updateApplicant(applicant.id, {
                        visaRequirementChecked: true,
                        visaSupportRequired: e.target.checked,
                        visaDocumentsSent: e.target.checked ? applicant.visaDocumentsSent : false
                      })
                    }
                  />
                  <span>ビザ支援書類が必要</span>
                  <strong>{applicant.visaSupportRequired ? '必要' : applicant.visaRequirementChecked ? '不要' : '未確認'}</strong>
                </label>
                <label className="crm-check-card">
                  <input
                    type="checkbox"
                    checked={applicant.visaDocumentsSent}
                    onChange={(e) =>
                      updateApplicant(applicant.id, {
                        visaRequirementChecked: e.target.checked ? true : applicant.visaRequirementChecked,
                        visaSupportRequired: e.target.checked ? true : applicant.visaSupportRequired,
                        visaDocumentsSent: e.target.checked,
                        visaDocumentsSentDate: e.target.checked && !applicant.visaDocumentsSentDate ? getTodayInputValue() : applicant.visaDocumentsSentDate,
                        visaDocumentsSentStaff:
                          e.target.checked && !applicant.visaDocumentsSentStaff
                            ? operatorName.trim() || applicant.staff || applicant.reminderStaff
                            : applicant.visaDocumentsSentStaff
                      })
                    }
                  />
                  <span>ビザ支援書類送付済み</span>
                  <strong>{applicant.visaDocumentsSent ? '送付済み' : '未送付'}</strong>
                </label>
                <label>
                  送付日
                  <input
                    type="date"
                    value={applicant.visaDocumentsSentDate}
                    onChange={(e) => updateApplicant(applicant.id, { visaDocumentsSentDate: e.target.value })}
                  />
                </label>
                <label>
                  送付担当者
                  <input
                    type="text"
                    placeholder="例：Tanaka"
                    value={applicant.visaDocumentsSentStaff}
                    onChange={(e) => updateApplicant(applicant.id, { visaDocumentsSentStaff: e.target.value })}
                  />
                </label>
              </div>
              <div className="visa-status-row">
                <span className={getVisaStatusClassName(applicant)}>{getVisaStatusLabel(applicant)}</span>
                <button type="button" onClick={() => markVisaDocumentsAsSentForApplicant(applicant)}>
                  ビザ支援書類を送付済みにする
                </button>
              </div>
            </section>

            <section className="crm-section-card print-section-confirmations">
              <h3>確認事項・食事制限</h3>
              <div className="crm-status-board">
                {confirmationDefinitions.map((definition) => (
                  <label key={definition.key} className="crm-check-card">
                    <input
                      type="checkbox"
                      checked={applicant[definition.key]}
                      onChange={(e) =>
                        updateApplicant(applicant.id, { [definition.key]: e.target.checked } as Partial<Applicant>)
                      }
                    />
                    <span>{definition.label}</span>
                    <strong>{applicant[definition.key] ? '済' : '未確認'}</strong>
                  </label>
                ))}
              </div>

              <div className="dietary-detail-grid">
                <section className="dietary-detail-card">
                  <h4>食物アレルギー</h4>
                  <label className="inline-check">
                    <input
                      type="checkbox"
                      checked={applicant.hasFoodAllergy}
                      onChange={(e) => updateApplicant(applicant.id, { hasFoodAllergy: e.target.checked })}
                    />
                    <span>食物アレルギーあり</span>
                  </label>
                  {applicant.hasFoodAllergy ? (
                    <>
                      <div className="checkbox-group">
                        {allergyOptions.map((option) => (
                          <label key={option.key} className="inline-check">
                            <input
                              type="checkbox"
                              checked={applicant[option.key]}
                              onChange={(e) =>
                                updateApplicant(applicant.id, { [option.key]: e.target.checked } as Partial<Applicant>)
                              }
                            />
                            <span>{option.label}</span>
                          </label>
                        ))}
                      </div>
                      <label>
                        その他・詳細
                        <textarea
                          rows={3}
                          placeholder="例：キウイ、マンゴー、重度の場合の注意点など"
                          value={applicant.allergyOther}
                          onChange={(e) => updateApplicant(applicant.id, { allergyOther: e.target.value })}
                        />
                      </label>
                    </>
                  ) : (
                    <p className="empty-note">なしの場合はチェックを外し、確認後に「食物アレルギー・宗教的食物制限確認」を済にしてください。</p>
                  )}
                </section>

                <section className="dietary-detail-card">
                  <h4>宗教的な食物制限</h4>
                  <label className="inline-check">
                    <input
                      type="checkbox"
                      checked={applicant.hasReligiousDietaryRestriction}
                      onChange={(e) =>
                        updateApplicant(applicant.id, { hasReligiousDietaryRestriction: e.target.checked })
                      }
                    />
                    <span>宗教的に避ける食べものあり</span>
                  </label>
                  {applicant.hasReligiousDietaryRestriction ? (
                    <>
                      <div className="checkbox-group">
                        {religiousDietaryOptions.map((option) => (
                          <label key={option.key} className="inline-check">
                            <input
                              type="checkbox"
                              checked={applicant[option.key]}
                              onChange={(e) =>
                                updateApplicant(applicant.id, { [option.key]: e.target.checked } as Partial<Applicant>)
                              }
                            />
                            <span>{option.label}</span>
                          </label>
                        ))}
                      </div>
                      <label>
                        その他・詳細
                        <textarea
                          rows={3}
                          placeholder="例：調味料のアルコール不可、同じ調理器具不可など"
                          value={applicant.religiousOther}
                          onChange={(e) => updateApplicant(applicant.id, { religiousOther: e.target.value })}
                        />
                      </label>
                    </>
                  ) : (
                    <p className="empty-note">なしの場合はチェックを外し、確認後に「食物アレルギー・宗教的食物制限確認」を済にしてください。</p>
                  )}
                </section>
              </div>
            </section>

            <section className="crm-section-card print-section-communication">
              <h3>やり取り履歴・メール記録</h3>
              {hasAnyCommunication ? (
                <div className="crm-timeline">
                  {applicant.communicationHistory.trim() && (
                    <article className="crm-timeline-item crm-timeline-primary">
                      <span>時系列メモ・対応履歴</span>
                      <pre>{applicant.communicationHistory}</pre>
                    </article>
                  )}
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

            <section className="crm-section-card print-section-add-history">
              <h3>対応履歴を追加</h3>
              <p className="empty-note">学生からの連絡、こちらから送ったメール、電話、書類確認などをここに残せます。</p>
              <div className="communication-form-grid">
                <label>
                  日付
                  <input
                    type="date"
                    value={communicationDraft.date}
                    onChange={(e) => updateCommunicationDraft({ date: e.target.value })}
                  />
                </label>
                <label>
                  種別
                  <select
                    value={communicationDraft.type}
                    onChange={(e) => updateCommunicationDraft({ type: e.target.value as CommunicationType })}
                  >
                    {communicationTypeOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  担当者
                  <input
                    type="text"
                    placeholder="例：田中"
                    value={communicationDraft.staff}
                    onChange={(e) => updateCommunicationDraft({ staff: e.target.value })}
                  />
                </label>
                <label>
                  件名・見出し
                  <input
                    type="text"
                    placeholder="例：OneDrive linkについて問い合わせ"
                    value={communicationDraft.subject}
                    onChange={(e) => updateCommunicationDraft({ subject: e.target.value })}
                  />
                </label>
                <label>
                  要約
                  <textarea
                    rows={3}
                    placeholder="例：学生よりOneDriveリンクが開けないとの連絡あり。新しいリンクを案内する必要あり。"
                    value={communicationDraft.summary}
                    onChange={(e) => updateCommunicationDraft({ summary: e.target.value })}
                  />
                </label>
                <label>
                  本文・メモ貼付
                  <textarea
                    rows={7}
                    placeholder="メール本文や電話メモを貼り付け"
                    value={communicationDraft.body}
                    onChange={(e) => updateCommunicationDraft({ body: e.target.value })}
                  />
                </label>
                <label>
                  次にやることへ入れる内容
                  <textarea
                    rows={3}
                    placeholder="例：新しいOneDriveリンクを送る。"
                    value={communicationDraft.nextAction}
                    onChange={(e) => updateCommunicationDraft({ nextAction: e.target.value })}
                  />
                </label>
                <label className="communication-check">
                  <input
                    type="checkbox"
                    checked={communicationDraft.reflectNextAction}
                    onChange={(e) => updateCommunicationDraft({ reflectNextAction: e.target.checked })}
                  />
                  <span>「次にやること」へ反映する</span>
                </label>
              </div>
              <button type="button" className="primary-action" onClick={() => addCommunicationRecord(applicant)}>
                この学生の対応履歴に追加
              </button>
            </section>

            <section className="crm-section-card print-section-edit-history">
              <h3>記録・修正</h3>
              <div className="crm-form-grid">
                <label>
                  やり取り履歴
                  <textarea
                    rows={6}
                    value={applicant.communicationHistory}
                    onChange={(e) => updateApplicant(applicant.id, { communicationHistory: e.target.value })}
                  />
                </label>
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
            <section className="crm-section-card crm-next-action-card print-section-next-action">
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

            <section className="crm-section-card print-section-memo">
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

            <section className="crm-section-card print-section-staff">
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
              <section className="crm-section-card print-section-reminder">
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
                      placeholder="例：田中（空欄なら作業者名）"
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

        <label>
          作業者名:
          <input
            type="text"
            placeholder="例：Tanaka"
            value={operatorName}
            onChange={(e) => setOperatorName(e.target.value)}
          />
        </label>

        <span className="export-file-name-preview">出力名：{exportFileNamePreview}</span>

        <button onClick={applyDueDate} disabled={!applicants.length || !dueDate}>
          全員に期日を適用
        </button>
        <button onClick={applyDueDateToUnsetOnly} disabled={!applicants.length || !dueDate || unsetDueDateCount === 0}>
          期限未設定者のみに適用
        </button>
        <button onClick={exportStatus} disabled={!applicants.length}>
          進捗Excelを出力
        </button>
        <button type="button" onClick={printApplicantList} disabled={!applicants.length}>
          一覧を印刷
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
          <div className="summary-card summary-unset">
            <span>確認事項未完了</span>
            <strong>{confirmationPendingCount}名</strong>
          </div>
          <div className="summary-card summary-unset">
            <span>ビザ要否未確認</span>
            <strong>{visaRequirementUncheckedCount}名</strong>
          </div>
          <div className="summary-card summary-soon">
            <span>ビザ未送付</span>
            <strong>{visaDocumentsPendingCount}名</strong>
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

      {!!applicants.length && dietaryAttentionApplicants.length > 0 && (
        <section className="dietary-summary-panel">
          <h2>食物アレルギー・宗教的食物制限 確認リスト</h2>
          <p>食事手配や訪問先共有が必要になりそうな応募者です。</p>
          <div className="dietary-list">
            {dietaryAttentionApplicants.map((applicant) => (
              <button
                key={applicant.id}
                type="button"
                className="dietary-list-item"
                onClick={() => openApplicantPage(applicant.id)}
              >
                <strong>{applicant.name || '氏名未入力'}</strong>
                <span>アレルギー：{getFoodAllergySummary(applicant)}</span>
                <span>宗教的制限：{getReligiousDietarySummary(applicant)}</span>
              </button>
            ))}
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
                <th>ビザ</th>
                <th>パスポートコピー</th>
                <th>在籍証明書</th>
                <th>追加書類</th>
                <th>確認事項</th>
                <th>食事制限</th>
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
                      <span className={getVisaStatusClassName(a)}>{getVisaStatusLabel(a)}</span>
                    </td>
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
                    <td>{getDocumentCompletionLabel(a)}</td>
                    <td>{getConfirmationCompletionLabel(a)}</td>
                    <td>
                      アレルギー：{getFoodAllergySummary(a)}
                      <br />
                      宗教食：{getReligiousDietarySummary(a)}
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
                  <td colSpan={16}>該当する応募者はいません。</td>
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