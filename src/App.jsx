import React, { useState, useEffect } from 'react';
import { initializeApp } from 'firebase/app';
import { 
  getFirestore, 
  collection, 
  addDoc, 
  onSnapshot, 
  query, 
  orderBy, 
  deleteDoc, 
  doc 
} from 'firebase/firestore';
import { 
  Phone, 
  Tv, 
  Bell, 
  Clock, 
  CheckCircle2, 
  PlusCircle, 
  BarChart3, 
  ListFilter, 
  User, 
  Trash2, 
  Sparkles,
  Layers,
  FileCheck2,
  ChevronDown,
  Download
} from 'lucide-react';
import * as XLSX from 'xlsx';

// Конфигурация Firebase
const firebaseConfig = {
  apiKey: "AIzaSyCB-Xg1gRVb0AEGW2BTlhAmzTa4xWK03SA",
  authDomain: "bmc-comm-sector.firebaseapp.com",
  projectId: "bmc-comm-sector",
  storageBucket: "bmc-comm-sector.firebasestorage.app",
  messagingSenderId: "454392409266",
  appId: "1:454392409266:web:aa211ed091c74283c36dab",
  measurementId: "G-XY67HHX3B3"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Список постоянных сотрудников
const PERFORMERS = ['Роман', 'Данияр', 'Нурыш', 'Максат'];

const CATEGORIES = [
  { id: 'phone', label: 'Телефоны', icon: Phone, color: 'from-blue-600 to-indigo-600', textColor: 'text-blue-400', badgeBg: 'bg-blue-500/10 border-blue-500/30' },
  { id: 'tv', label: 'Телевизоры', icon: Tv, color: 'from-purple-600 to-pink-600', textColor: 'text-purple-400', badgeBg: 'bg-purple-500/10 border-purple-500/30' },
  { id: 'nurse_call', label: 'Палатная связь', icon: Bell, color: 'from-amber-600 to-orange-600', textColor: 'text-amber-400', badgeBg: 'bg-amber-500/10 border-amber-500/30' },
  { id: 'clock', label: 'Часофикация', icon: Clock, color: 'from-emerald-600 to-teal-600', textColor: 'text-emerald-400', badgeBg: 'bg-emerald-500/10 border-emerald-500/30' }
];

export default function App() {
  const [activeTab, setActiveTab] = useState('add');
  const [tickets, setTickets] = useState([]);
  const [loading, setLoading] = useState(true);

  // Форма добавления
  const [performer, setPerformer] = useState(PERFORMERS[0]);
  const [category, setCategory] = useState(CATEGORIES[0].id);
  const [description, setDescription] = useState('');
  const [isAspans, setIsAspans] = useState(true);
  const [successMsg, setSuccessMsg] = useState(false);

  // Фильтры отчета
  const [reportPeriod, setReportPeriod] = useState('month');

  // Подписка на данные Firebase
  useEffect(() => {
    const q = query(collection(db, 'tickets'), orderBy('createdAt', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const ticketsData = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setTickets(ticketsData);
      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!description.trim()) return;

    try {
      await addDoc(collection(db, 'tickets'), {
        performer,
        category,
        description: description.trim(),
        isAspans,
        createdAt: new Date().toISOString()
      });

      setDescription('');
      setSuccessMsg(true);
      setTimeout(() => setSuccessMsg(false), 2500);
    } catch (err) {
      console.error("Ошибка при сохранении:", err);
    }
  };

  const handleDelete = async (id) => {
    if (confirm('Удалить эту заявку?')) {
      await deleteDoc(doc(db, 'tickets', id));
    }
  };

  // Расчет отчетов
  const filterTicketsByPeriod = () => {
    const now = new Date();
    return tickets.filter(t => {
      const date = new Date(t.createdAt);
      if (reportPeriod === 'month') {
        return date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
      } else {
        const currentQuarter = Math.floor(now.getMonth() / 3);
        const ticketQuarter = Math.floor(date.getMonth() / 3);
        return currentQuarter === ticketQuarter && date.getFullYear() === now.getFullYear();
      }
    });
  };

  const filteredTickets = filterTicketsByPeriod();

  // Функция экспорта в Excel
  const exportToExcel = () => {
    if (filteredTickets.length === 0) {
      alert('За выбранный период нет заявок для выгрузки!');
      return;
    }

    const excelData = filteredTickets.map((t, index) => ({
      "№": index + 1,
      "Дата": new Date(t.createdAt).toLocaleDateString('ru-RU'),
      "Время": new Date(t.createdAt).toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' }),
      "Исполнитель": t.performer || 'Не указан',
      "Источник": t.isAspans !== false ? "Через Aspans" : "Заявка вне Aspans",
      "Тип системы": CATEGORIES.find(c => c.id === t.category)?.label || t.category,
      "Описание работ": t.description
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    const sheetName = reportPeriod === 'month' ? "Заявки_Месяц" : "Заявки_Квартал";
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    const fileName = `Отчет_Заявки_${reportPeriod === 'month' ? 'Месяц' : 'Квартал'}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // Статистика по сотрудникам
  const performerStats = filteredTickets.reduce((acc, ticket) => {
    const perf = ticket.performer || 'Не указан';
    if (!acc[perf]) {
      acc[perf] = { total: 0, aspans: 0, nonAspans: 0, byCat: {} };
    }
    acc[perf].total += 1;
    
    const isAsp = ticket.isAspans !== false;
    if (isAsp) {
      acc[perf].aspans += 1;
    } else {
      acc[perf].nonAspans += 1;
    }

    acc[perf].byCat[ticket.category] = (acc[perf].byCat[ticket.category] || 0) + 1;
    return acc;
  }, {});

  // Статистика по категориям
  const categoryStats = filteredTickets.reduce((acc, ticket) => {
    acc[ticket.category] = (acc[ticket.category] || 0) + 1;
    return acc;
  }, {});

  // Статистика по источнику
  const aspansStats = filteredTickets.reduce((acc, ticket) => {
    const isAsp = ticket.isAspans !== false;
    if (isAsp) acc.aspans += 1;
    else acc.nonAspans += 1;
    return acc;
  }, { aspans: 0, nonAspans: 0 });

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans pb-24">
      {/* Шапка */}
      <header className="sticky top-0 z-50 bg-slate-900/80 backdrop-blur-md border-b border-slate-800 px-4 py-3">
        <div className="max-w-md mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-gradient-to-tr from-cyan-500 to-blue-600 p-2 rounded-xl text-white shadow-lg shadow-cyan-500/20">
              <Layers className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-bold text-base leading-tight">Учет Заявок</h1>
              <p className="text-xs text-slate-400">BMC Comm Sector</p>
            </div>
          </div>
          <span className="text-xs bg-slate-800 text-cyan-400 px-2.5 py-1 rounded-full border border-slate-700 font-medium">
            Онлайн
          </span>
        </div>
      </header>

      <main className="max-w-md mx-auto px-4 pt-4">
        {/* Вкладка 1: Добавление */}
        {activeTab === 'add' && (
          <div className="space-y-4">
            {successMsg && (
              <div className="p-3 bg-emerald-500/10 border border-emerald-500/30 rounded-2xl flex items-center gap-3 text-emerald-400 text-sm animate-fade-in">
                <CheckCircle2 className="w-5 h-5 flex-shrink-0" />
                <span>Заявка успешно сохранена!</span>
              </div>
            )}

            <form onSubmit={handleSubmit} className="bg-slate-900 border border-slate-800 rounded-3xl p-4 shadow-xl space-y-4">
              {/* Исполнитель */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 mb-1.5 uppercase tracking-wider">
                  Исполнитель
                </label>
                <div className="relative">
                  <User className="w-5 h-5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 pointer-events-none z-10" />
                  <select
                    value={performer}
                    onChange={(e) => setPerformer(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-xl py-2.5 pl-10 pr-10 text-sm focus:outline-none focus:border-cyan-500 text-slate-100 transition-colors appearance-none cursor-pointer font-medium"
                  >
                    {PERFORMERS.map((name) => (
                      <option key={name} value={name} className="bg-slate-900 text-slate-100">
                        {name}
                      </option>
                    ))}
                  </select>
                  <ChevronDown className="w-4 h-4 absolute right-3 top-1/2 -translate-y-1/2 text-slate-500 pointer-events-none" />
                </div>
              </div>

              {/* Источник заявки */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 mb-1.5 uppercase tracking-wider">
                  Источник заявки
                </label>
                <div className="grid grid-cols-2 gap-2 p-1 bg-slate-950 border border-slate-800 rounded-2xl">
                  <button
                    type="button"
                    onClick={() => setIsAspans(true)}
                    className={`py-2 px-3 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 transition-all ${
                      isAspans 
                        ? 'bg-cyan-500 text-slate-950 shadow-md shadow-cyan-500/20' 
                        : 'text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    Через Aspans
                  </button>
                  <button
                    type="button"
                    onClick={() => setIsAspans(false)}
                    className={`py-2 px-3 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 transition-all ${
                      !isAspans 
                        ? 'bg-slate-800 text-amber-400 border border-amber-500/30' 
                        : 'text-slate-400 hover:text-slate-200'
                    }`}
                  >
                    <FileCheck2 className="w-3.5 h-3.5" />
                    Заявка вне Aspans
                  </button>
                </div>
              </div>

              {/* Тип системы */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 mb-1.5 uppercase tracking-wider">
                  Тип системы
                </label>
                <div className="grid grid-cols-2 gap-2">
                  {CATEGORIES.map((cat) => {
                    const Icon = cat.icon;
                    const isSelected = category === cat.id;
                    return (
                      <button
                        key={cat.id}
                        type="button"
                        onClick={() => setCategory(cat.id)}
                        className={`p-3 rounded-2xl border text-left flex flex-col justify-between h-20 transition-all ${
                          isSelected
                            ? `bg-gradient-to-br ${cat.color} border-transparent text-white shadow-lg`
                            : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                        }`}
                      >
                        <Icon className={`w-5 h-5 ${isSelected ? 'text-white' : cat.textColor}`} />
                        <span className="text-xs font-semibold">{cat.label}</span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Описание */}
              <div>
                <label className="block text-xs font-semibold text-slate-400 mb-1.5 uppercase tracking-wider">
                  Описание выполненных работ
                </label>
                <textarea
                  required
                  rows={3}
                  placeholder="Например: Терапия 3, палата 5. Замена приставки..."
                  value={description}
                  onChange={(e) => setDescription(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm focus:outline-none focus:border-cyan-500 text-slate-100 placeholder:text-slate-600 transition-colors resize-none"
                />
              </div>

              {/* Кнопка отправки */}
              <button
                type="submit"
                className="w-full bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-slate-950 font-bold py-3.5 px-4 rounded-2xl shadow-lg shadow-cyan-500/25 flex items-center justify-center gap-2 text-sm transition-all active:scale-[0.98]"
              >
                <PlusCircle className="w-5 h-5" />
                Сохранить заявку
              </button>
            </form>
          </div>
        )}

        {/* Вкладка 2: Журнал */}
        {activeTab === 'list' && (
          <div className="space-y-3">
            <div className="flex items-center justify-between px-1">
              <h2 className="font-bold text-sm text-slate-300">Всего заявок: {tickets.length}</h2>
            </div>

            {loading ? (
              <div className="text-center py-12 text-slate-500 text-sm">Загрузка данных...</div>
            ) : tickets.length === 0 ? (
              <div className="bg-slate-900 border border-slate-800 rounded-3xl p-8 text-center text-slate-500 text-sm">
                Заявок пока нет
              </div>
            ) : (
              tickets.map((ticket) => {
                const catObj = CATEGORIES.find(c => c.id === ticket.category) || CATEGORIES[0];
                const Icon = catObj.icon;
                const isAsp = ticket.isAspans !== false;

                return (
                  <div key={ticket.id} className="bg-slate-900 border border-slate-800 rounded-2xl p-3.5 space-y-2 relative group">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className={`p-1.5 rounded-lg border text-xs font-medium flex items-center gap-1 ${catObj.badgeBg} ${catObj.textColor}`}>
                          <Icon className="w-3.5 h-3.5" />
                          {catObj.label}
                        </span>
                        
                        <span className={`text-[10px] px-2 py-0.5 rounded-full border ${
                          isAsp 
                            ? 'bg-cyan-500/10 border-cyan-500/30 text-cyan-400' 
                            : 'bg-amber-500/10 border-amber-500/30 text-amber-400'
                        }`}>
                          {isAsp ? 'Aspans' : 'Вне Aspans'}
                        </span>
                      </div>

                      <div className="flex items-center gap-2">
                        <span className="text-[11px] text-slate-500">
                          {new Date(ticket.createdAt).toLocaleDateString('ru-RU', { day: 'numeric', month: 'short' })}
                        </span>
                        <button 
                          onClick={() => handleDelete(ticket.id)}
                          className="text-slate-600 hover:text-red-400 p-1 rounded-lg transition-colors"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>

                    <p className="text-xs text-slate-200 font-medium leading-relaxed">
                      {ticket.description}
                    </p>

                    <div className="flex items-center gap-1 text-[11px] text-slate-400 pt-1 border-t border-slate-800/60">
                      <User className="w-3 h-3 text-slate-500" />
                      <span className="font-semibold text-slate-300">{ticket.performer}</span>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        )}

        {/* Вкладка 3: Отчеты */}
        {activeTab === 'report' && (
          <div className="space-y-4">
            {/* Выбор периода */}
            <div className="flex items-center justify-between bg-slate-900 border border-slate-800 p-1.5 rounded-2xl">
              <button
                onClick={() => setReportPeriod('month')}
                className={`flex-1 py-1.5 rounded-xl text-xs font-semibold transition-all ${
                  reportPeriod === 'month' ? 'bg-slate-800 text-cyan-400 shadow-sm' : 'text-slate-400'
                }`}
              >
                Текущий месяц
              </button>
              <button
                onClick={() => setReportPeriod('quarter')}
                className={`flex-1 py-1.5 rounded-xl text-xs font-semibold transition-all ${
                  reportPeriod === 'quarter' ? 'bg-slate-800 text-cyan-400 shadow-sm' : 'text-slate-400'
                }`}
              >
                Текущий квартал
              </button>
            </div>

            {/* Кнопка экспорта в Excel */}
            <button
              onClick={exportToExcel}
              className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-3 px-4 rounded-2xl shadow-lg shadow-emerald-600/20 flex items-center justify-center gap-2 text-xs transition-all active:scale-[0.98]"
            >
              <Download className="w-4 h-4" />
              Скачать отчет в Excel ({reportPeriod === 'month' ? 'Месяц' : 'Квартал'})
            </button>

            {/* Карточка общей сводки */}
            <div className="bg-slate-900 border border-slate-800 rounded-3xl p-4 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Всего выполнено</span>
                <span className="text-2xl font-black text-cyan-400">{filteredTickets.length}</span>
              </div>

              {/* Разделение Aspans vs Вне Aspans */}
              <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-800">
                <div className="bg-slate-950 p-2.5 rounded-xl border border-slate-800">
                  <div className="text-[10px] text-slate-400 mb-0.5">Через Aspans</div>
                  <div className="text-base font-bold text-cyan-400">{aspansStats.aspans}</div>
                </div>
                <div className="bg-slate-950 p-2.5 rounded-xl border border-slate-800">
                  <div className="text-[10px] text-slate-400 mb-0.5">Вне Aspans</div>
                  <div className="text-base font-bold text-amber-400">{aspansStats.nonAspans}</div>
                </div>
              </div>
            </div>

            {/* По категориям */}
            <div className="bg-slate-900 border border-slate-800 rounded-3xl p-4 space-y-3">
              <h3 className="text-xs font-semibold text-slate-400 uppercase tracking-wider">По типам систем</h3>
              <div className="space-y-2">
                {CATEGORIES.map(cat => {
                  const count = categoryStats[cat.id] || 0;
                  const percent = filteredTickets.length ? Math.round((count / filteredTickets.length) * 100) : 0;
                  const Icon = cat.icon;
                  return (
                    <div key={cat.id} className="space-y-1">
                      <div className="flex items-center justify-between text-xs">
                        <span className="flex items-center gap-1.5 text-slate-300">
                          <Icon className={`w-3.5 h-3.5 ${cat.textColor}`} />
                          {cat.label}
                        </span>
                        <span className="font-semibold text-slate-200">{count} ({percent}%)</span>
                      </div>
                      <div className="h-1.5 bg-slate-950 rounded-full overflow-hidden">
                        <div 
                          className={`h-full bg-gradient-to-r ${cat.color}`} 
                          style={{ width: `${percent}%` }}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* По сотрудникам */}
            <div className="bg-slate-900 border border-slate-800 rounded-3xl p-4 space-y-3">
              <h3 className="text-xs font-semibold text-slate-400 uppercase tracking-wider">По сотрудникам</h3>
              <div className="space-y-3">
                {PERFORMERS.map((name) => {
                  const stat = performerStats[name] || { total: 0, aspans: 0, nonAspans: 0, byCat: {} };
                  return (
                    <div key={name} className="bg-slate-950 border border-slate-800 rounded-2xl p-3 space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-xs text-slate-200">{name}</span>
                        <span className="text-xs font-bold text-cyan-400">{stat.total} заявок</span>
                      </div>
                      
                      <div className="flex gap-2 text-[10px]">
                        <span className="text-cyan-400/90 bg-cyan-500/10 px-2 py-0.5 rounded-md border border-cyan-500/20">
                          Aspans: {stat.aspans}
                        </span>
                        <span className="text-amber-400/90 bg-amber-500/10 px-2 py-0.5 rounded-md border border-amber-500/20">
                          Вне Aspans: {stat.nonAspans}
                        </span>
                      </div>

                      <div className="flex flex-wrap gap-1.5 pt-1">
                        {CATEGORIES.map(cat => {
                          const catCount = stat.byCat[cat.id];
                          if (!catCount) return null;
                          return (
                            <span key={cat.id} className="text-[10px] text-slate-400 bg-slate-900 px-2 py-0.5 rounded-md">
                              {cat.label}: {catCount}
                            </span>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Нижняя навигация */}
      <nav className="fixed bottom-0 left-0 right-0 bg-slate-900/90 backdrop-blur-lg border-t border-slate-800 px-6 py-2 z-50">
        <div className="max-w-md mx-auto flex items-center justify-around">
          <button
            onClick={() => setActiveTab('add')}
            className={`flex flex-col items-center gap-1 py-1 transition-all ${
              activeTab === 'add' ? 'text-cyan-400' : 'text-slate-500 hover:text-slate-400'
            }`}
          >
            <PlusCircle className="w-5 h-5" />
            <span className="text-[10px] font-medium">Новая</span>
          </button>

          <button
            onClick={() => setActiveTab('list')}
            className={`flex flex-col items-center gap-1 py-1 transition-all ${
              activeTab === 'list' ? 'text-cyan-400' : 'text-slate-500 hover:text-slate-400'
            }`}
          >
            <ListFilter className="w-5 h-5" />
            <span className="text-[10px] font-medium">Журнал</span>
          </button>

          <button
            onClick={() => setActiveTab('report')}
            className={`flex flex-col items-center gap-1 py-1 transition-all ${
              activeTab === 'report' ? 'text-cyan-400' : 'text-slate-500 hover:text-slate-400'
            }`}
          >
            <BarChart3 className="w-5 h-5" />
            <span className="text-[10px] font-medium">Отчеты</span>
          </button>
        </div>
      </nav>
    </div>
  );
}
