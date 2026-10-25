import type { Lang } from '@genoffice/i18n'

export interface OfficeActionLabels {
  open: string
  save: string
  saveAs: string
  chat: string
  chatSelection: string
}

export const OFFICE_ACTION_LABELS: Record<Lang, OfficeActionLabels> = {
  zh: {
    open: '打开…',
    save: '保存',
    saveAs: '另存为…',
    chat: 'VS Code 聊天',
    chatSelection: '询问选区',
  },
  en: {
    open: 'Open…',
    save: 'Save',
    saveAs: 'Save As…',
    chat: 'VS Code Chat',
    chatSelection: 'Ask about selection',
  },
  ja: {
    open: '開く…',
    save: '保存',
    saveAs: '名前を付けて保存…',
    chat: 'VS Code チャット',
    chatSelection: '選択範囲について質問',
  },
  ko: {
    open: '열기…',
    save: '저장',
    saveAs: '다른 이름으로 저장…',
    chat: 'VS Code 채팅',
    chatSelection: '선택 영역에 대해 질문',
  },
  fr: {
    open: 'Ouvrir…',
    save: 'Enregistrer',
    saveAs: 'Enregistrer sous…',
    chat: 'Chat VS Code',
    chatSelection: 'Interroger la sélection',
  },
  de: {
    open: 'Öffnen…',
    save: 'Speichern',
    saveAs: 'Speichern unter…',
    chat: 'VS Code Chat',
    chatSelection: 'Zur Auswahl fragen',
  },
  es: {
    open: 'Abrir…',
    save: 'Guardar',
    saveAs: 'Guardar como…',
    chat: 'Chat de VS Code',
    chatSelection: 'Preguntar sobre la selección',
  },
  th: {
    open: 'เปิด…',
    save: 'บันทึก',
    saveAs: 'บันทึกเป็น…',
    chat: 'แชต VS Code',
    chatSelection: 'ถามเกี่ยวกับส่วนที่เลือก',
  },
  id: {
    open: 'Buka…',
    save: 'Simpan',
    saveAs: 'Simpan sebagai…',
    chat: 'Chat VS Code',
    chatSelection: 'Tanyakan tentang pilihan',
  },
  ru: {
    open: 'Открыть…',
    save: 'Сохранить',
    saveAs: 'Сохранить как…',
    chat: 'Чат VS Code',
    chatSelection: 'Спросить о выделении',
  },
  ar: {
    open: 'فتح…',
    save: 'حفظ',
    saveAs: 'حفظ باسم…',
    chat: 'دردشة VS Code',
    chatSelection: 'السؤال عن التحديد',
  },
  pt: {
    open: 'Abrir…',
    save: 'Salvar',
    saveAs: 'Salvar como…',
    chat: 'Chat do VS Code',
    chatSelection: 'Perguntar sobre a seleção',
  },
  it: {
    open: 'Apri…',
    save: 'Salva',
    saveAs: 'Salva con nome…',
    chat: 'Chat di VS Code',
    chatSelection: 'Chiedi della selezione',
  },
  pl: {
    open: 'Otwórz…',
    save: 'Zapisz',
    saveAs: 'Zapisz jako…',
    chat: 'Czat VS Code',
    chatSelection: 'Zapytaj o zaznaczenie',
  },
  cs: {
    open: 'Otevřít…',
    save: 'Uložit',
    saveAs: 'Uložit jako…',
    chat: 'Chat VS Code',
    chatSelection: 'Zeptat se na výběr',
  },
  nl: {
    open: 'Openen…',
    save: 'Opslaan',
    saveAs: 'Opslaan als…',
    chat: 'VS Code Chat',
    chatSelection: 'Vragen over selectie',
  },
  ms: {
    open: 'Buka…',
    save: 'Simpan',
    saveAs: 'Simpan sebagai…',
    chat: 'Sembang VS Code',
    chatSelection: 'Tanya tentang pilihan',
  },
  he: {
    open: 'פתיחה…',
    save: 'שמירה',
    saveAs: 'שמירה בשם…',
    chat: 'צ׳אט VS Code',
    chatSelection: 'שאלה על הבחירה',
  },
  hi: {
    open: 'खोलें…',
    save: 'सहेजें',
    saveAs: 'इस रूप में सहेजें…',
    chat: 'VS Code चैट',
    chatSelection: 'चयन के बारे में पूछें',
  },
  'zh-TW': {
    open: '開啟…',
    save: '儲存',
    saveAs: '另存新檔…',
    chat: 'VS Code 聊天',
    chatSelection: '詢問選取範圍',
  },
  vi: {
    open: 'Mở…',
    save: 'Lưu',
    saveAs: 'Lưu dưới dạng…',
    chat: 'Trò chuyện VS Code',
    chatSelection: 'Hỏi về vùng chọn',
  },
}
