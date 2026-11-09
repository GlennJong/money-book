import { useState, useMemo } from 'react';
import { type Transaction } from '../../types';
import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip } from 'recharts';

type DiscoverScreenProps = {
  transactions: Transaction[];
};

type Unit = 'month' | 'quarter' | 'half-year' | 'year';

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884d8', '#82ca9d', '#ffc658', '#8dd1e1', '#a4de6c', '#d0ed57'];

const DiscoverScreen = ({ transactions }: DiscoverScreenProps) => {
  const [unit, setUnit] = useState<Unit>('month');

  // Load threshold from config
  const bigPurchaseThreshold = useMemo(() => {
    const saved = localStorage.getItem('moneybook_big_purchase_threshold');
    return saved ? Number(saved) : 1000;
  }, []);

  const reportData = useMemo(() => {
    const now = new Date();
    
    let startDate: Date;
    let endDate: Date;
    let prevStartDate: Date;
    let prevEndDate: Date;
    let label = '';

    if (unit === 'month') {
        startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        endDate = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
        label = startDate.toLocaleString('default', { month: 'long', year: 'numeric' });

        prevStartDate = new Date(now.getFullYear(), now.getMonth() - 2, 1);
        prevEndDate = new Date(now.getFullYear(), now.getMonth() - 1, 0, 23, 59, 59);
    } else if (unit === 'quarter') {
        const currentQuarter = Math.floor(now.getMonth() / 3);
        const startMonth = (currentQuarter - 1) * 3;
        startDate = new Date(now.getFullYear(), startMonth, 1);
        endDate = new Date(now.getFullYear(), startMonth + 3, 0, 23, 59, 59);
        const q = Math.floor(startDate.getMonth() / 3) + 1;
        label = `Q${q} ${startDate.getFullYear()}`;

        prevStartDate = new Date(now.getFullYear(), startMonth - 3, 1);
        prevEndDate = new Date(now.getFullYear(), startMonth, 0, 23, 59, 59);
    } else if (unit === 'half-year') {
        const currentHalf = Math.floor(now.getMonth() / 6);
        const startMonth = (currentHalf - 1) * 6;
        startDate = new Date(now.getFullYear(), startMonth, 1);
        endDate = new Date(now.getFullYear(), startMonth + 6, 0, 23, 59, 59);
        const h = Math.floor(startDate.getMonth() / 6) + 1;
        label = `${startDate.getFullYear()} H${h}`;
        
        prevStartDate = new Date(now.getFullYear(), startMonth - 6, 1);
        prevEndDate = new Date(now.getFullYear(), startMonth, 0, 23, 59, 59);
    } else { // year
        startDate = new Date(now.getFullYear() - 1, 0, 1);
        endDate = new Date(now.getFullYear() - 1, 11, 31, 23, 59, 59);
        label = startDate.getFullYear().toString();

        prevStartDate = new Date(now.getFullYear() - 2, 0, 1);
        prevEndDate = new Date(now.getFullYear() - 2, 11, 31, 23, 59, 59);
    }

    const currentPeriodTx = transactions.filter(t => {
        const d = new Date(t.created_at);
        return d >= startDate && d <= endDate;
    });

    const prevPeriodTx = transactions.filter(t => {
        const d = new Date(t.created_at);
        return d >= prevStartDate && d <= prevEndDate;
    });

    const currentTotal = currentPeriodTx.reduce((sum, t) => sum + t.price, 0);
    const prevTotal = prevPeriodTx.reduce((sum, t) => sum + t.price, 0);
    const diff = currentTotal - prevTotal;
    const diffPercent = prevTotal === 0 ? (currentTotal > 0 ? 100 : 0) : ((diff / prevTotal) * 100);

    const tagMap = new Map<string, number>();
    currentPeriodTx.forEach(t => {
        const validTags = t.tags?.filter(tag => tag && tag.trim() !== '') || [];
        const tag = validTags.length > 0 ? validTags[0] : 'unknown';
        tagMap.set(tag, (tagMap.get(tag) || 0) + t.price);
    });
    
    const pieData = Array.from(tagMap.entries())
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => b.value - a.value);

    const bigPurchases = [...currentPeriodTx]
        .filter(t => t.price >= bigPurchaseThreshold)
        .sort((a, b) => b.price - a.price);

    return {
        label,
        currentTotal,
        diff,
        diffPercent,
        pieData,
        bigPurchases
    };

  }, [transactions, unit, bigPurchaseThreshold]);

  return (
    <div style={{ padding: '20px', paddingBottom: '80px' }}>
      <h1>Discover Report</h1>
      
      <div style={{ display: 'flex', gap: '10px', marginBottom: '20px', overflowX: 'auto', paddingBottom: '5px' }}>
        {(['month', 'quarter', 'half-year', 'year'] as Unit[]).map((u) => (
            <button
                key={u}
                onClick={() => setUnit(u)}
                style={{
                    padding: '8px 16px',
                    borderRadius: '20px',
                    border: 'none',
                    backgroundColor: unit === u ? 'var(--primary)' : 'var(--input-bg)',
                    color: unit === u ? 'var(--text-inv)' : 'var(--text-main)',
                    cursor: 'pointer',
                    whiteSpace: 'nowrap',
                    textTransform: 'capitalize'
                }}
            >
                {u.replace('-', ' ')}
            </button>
        ))}
      </div>

      <div style={{ marginBottom: '25px', backgroundColor: 'var(--bg-card)', padding: '15px', borderRadius: '10px',  boxShadow: '0 2px 5px var(--shadow-color)' }}>
        <h2 style={{ margin: '0 0 10px 0', fontSize: '1.2em', color: 'var(--text-secondary)' }}>{reportData.label}</h2>
        <div style={{ fontSize: '2.5em', fontWeight: 'bold' }}>
            ${reportData.currentTotal.toLocaleString()}
        </div>
        <div style={{ color: reportData.diff > 0 ? 'var(--danger)' : 'var(--success)', fontWeight: 'bold' }}>
            {Math.abs(reportData.diff) > 0 ? (
                <>
                  {reportData.diff > 0 ? '↑' : '↓'} ${Math.abs(reportData.diff).toLocaleString()} ({Math.abs(reportData.diffPercent).toFixed(1)}%)
                </>
            ) : (
                <span>No change</span>
            )}
            <span style={{ color: 'var(--text-muted)', fontWeight: 'normal', fontSize: '0.8em', marginLeft: '5px' }}>vs prev. {unit}</span>
        </div>
      </div>

      {reportData.pieData.length > 0 ? (
          <div style={{ height: '300px', marginBottom: '30px' }}>
            <h3 style={{ marginBottom: '10px' }}>Spending by Tag</h3>
            <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                    <Pie
                        data={reportData.pieData}
                        cx="50%"
                        cy="50%"
                        labelLine={false}
                        outerRadius={100}
                        fill="#8884d8"
                        dataKey="value"
                        stroke="var(--bg-card)" // Add stroke to separate segments
                        label={({ x, y, cx, name, percent }) => (
                            <text 
                                x={x} 
                                y={y} 
                                fill="var(--text-secondary)" 
                                textAnchor={x > cx ? 'start' : 'end'} 
                                dominantBaseline="central" 
                                style={{ fontSize: '11px', fontWeight: '500' }}
                            >
                                {`${name} ${((percent ?? 0) * 100).toFixed(0)}%`}
                            </text>
                        )}
                    >
                        {reportData.pieData.map((_entry, index) => (
                            <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                    </Pie>
                    <Tooltip 
                        contentStyle={{ backgroundColor: 'var(--bg-card)', border: '1px solid var(--border-color)', borderRadius: '8px' }}
                        itemStyle={{ color: 'var(--text-main)' }}
                        formatter={(value) => value !== undefined ? `$${Number(value).toLocaleString()}` : '$0'} 
                    />
                    <Legend />
                </PieChart>
            </ResponsiveContainer>
          </div>
      ) : (
          <div style={{ padding: '20px', textAlign: 'center', color: 'var(--text-muted)' }}>
              No data for this period
          </div>
      )}

      <div>
        <h3 style={{ marginBottom: '15px' }}>Big Purchases (&#62; ${bigPurchaseThreshold.toLocaleString()})</h3>
        {reportData.bigPurchases.length > 0 ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {reportData.bigPurchases.map(t => (
                    <div key={t.id} style={{ 
                        display: 'flex', 
                        justifyContent: 'space-between', 
                        alignItems: 'center',
                        padding: '10px',
                        backgroundColor: 'var(--bg-item)',
                        borderRadius: '8px',
                        boxShadow: '0 1px 3px var(--shadow-color)'
                    }}>
                        <div>
                            <div style={{ fontWeight: 'bold' }}>{t.name}</div>
                            <div style={{ fontSize: '0.8em', color: 'var(--text-muted)' }}>
                                {new Date(t.created_at).toLocaleDateString()} • {t.tags?.filter(tag => tag && tag.trim() !== '').join(', ')}
                            </div>
                        </div>
                        <div style={{ fontWeight: 'bold', color: 'var(--danger)' }}>
                            ${t.price.toLocaleString()}
                        </div>
                    </div>
                ))}
            </div>
        ) : (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)' }}>No big purchases found</div>
        )}
      </div>
    </div>
  );
};

export default DiscoverScreen;
