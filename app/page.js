'use client';

import { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import {
  Chart as ChartJS,
  ArcElement,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend
} from 'chart.js';
import { Line, Pie, Bar } from 'react-chartjs-2';

ChartJS.register(
  ArcElement,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend
);

export default function Home() {
  const [data, setData] = useState(null);
  const [chartType, setChartType] = useState('line');
  const [headers, setHeaders] = useState([]);
  const [selectedX, setSelectedX] = useState('');
  const [selectedY, setSelectedY] = useState([]);
  const [fileName, setFileName] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterColumn, setFilterColumn] = useState('all');
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  const [darkMode, setDarkMode] = useState(false);
  const [chartTitle, setChartTitle] = useState('');
  const [customColors] = useState([
    '#3b82f6', '#10b981', '#f59e0b', 
    '#ef4444', '#8b5cf6', '#ec4899'
  ]);
  const [selectedXValues, setSelectedXValues] = useState([]);

  const fileInputRef = useRef(null);
  const chartRef = useRef(null);

  const handleFileUpload = (file) => {
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    const fileType = file.name.split('.').pop().toLowerCase();

    reader.onload = (event) => {
      try {
        let parsedData;

        if (fileType === 'csv') {
          Papa.parse(event.target.result, {
            header: true,
            complete: (results) => {
              parsedData = results.data.filter(row => 
                Object.values(row).some(val => val)
              );
              processData(parsedData);
            }
          });
        } else if (['xlsx', 'xls'].includes(fileType)) {
          const workbook = XLSX.read(event.target.result, { type: 'binary' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          parsedData = XLSX.utils.sheet_to_json(firstSheet);
          processData(parsedData);
        }
      } catch (error) {
        alert('文件解析失败: ' + error.message);
      }
    };

    if (fileType === 'csv') {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
  };

  const processData = (parsedData) => {
    if (parsedData.length === 0) return;

    const cols = Object.keys(parsedData[0]);
    setHeaders(cols);
    setSelectedX(cols[0] || '');
    setSelectedY([cols[1] || '']);
    setData(parsedData);
    setFilterColumn('all');
    setSearchTerm('');
    setSortConfig({ key: null, direction: 'asc' });
    setCurrentPage(1);
    setSelectedXValues([]);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFileUpload(file);
  };

  const xAxisUniqueValues = useMemo(() => {
    if (!data || !selectedX) return [];
    const values = [...new Set(data.map(row => String(row[selectedX])))];
    return values.sort();
  }, [data, selectedX]);

  const handleXValueToggle = (value) => {
    setSelectedXValues(prev =>
      prev.includes(value)
        ? prev.filter(v => v !== value)
        : [...prev, value]
    );
  };

  const handleSelectAllX = () => {
    setSelectedXValues(xAxisUniqueValues);
  };

  const handleClearAllX = () => {
    setSelectedXValues([]);
  };

  const chartData = useMemo(() => {
    if (!data || selectedY.length === 0) return null;

    // 如果有 X 轴筛选，只显示选中的值
    const filteredByX = selectedXValues.length > 0
      ? data.filter(row => selectedXValues.includes(String(row[selectedX])))
      : data;

    const labels = filteredByX.map(row => row[selectedX]);
    const datasets = selectedY.map((yCol, idx) => ({
      label: yCol,
      data: filteredByX.map(row => parseFloat(row[yCol]) || 0),
      backgroundColor: customColors[idx % customColors.length],
      borderColor: customColors[idx % customColors.length],
      borderWidth: 2,
      tension: 0.4,
    }));

    return { labels, datasets };
  }, [data, selectedX, selectedY, customColors, selectedXValues]);

  const stats = useMemo(() => {
    if (!data || selectedY.length === 0) return null;
    
    const allValues = selectedY.flatMap(col =>
      data.map(row => parseFloat(row[col])).filter(v => !isNaN(v))
    );

    if (allValues.length === 0) return null;

    const sum = allValues.reduce((a, b) => a + b, 0);
    const mean = sum / allValues.length;
    const sorted = [...allValues].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 === 0
      ? (sorted[mid - 1] + sorted[mid]) / 2
      : sorted[mid];

    return {
      mean: mean.toFixed(2),
      median: median.toFixed(2),
      max: Math.max(...allValues).toFixed(2),
      min: Math.min(...allValues).toFixed(2),
    };
  }, [data, selectedY]);

  const getChartOptions = (type) => {
    const baseOptions = {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        title: {
          display: !!chartTitle,
          text: chartTitle,
          font: { size: 18, weight: 'bold' },
          color: darkMode ? '#ffffff' : '#0a2540',
        },
        legend: {
          labels: {
            color: darkMode ? '#ffffff' : '#0a2540',
            font: { size: 14 },
          }
        },
        tooltip: {
          backgroundColor: darkMode ? '#1a1a1a' : '#ffffff',
          titleColor: darkMode ? '#ffffff' : '#0a2540',
          bodyColor: darkMode ? '#e0e0e0' : '#425466',
          borderColor: darkMode ? '#333' : '#e0e0e0',
          borderWidth: 1,
        }
      },
    };

    if (type === 'pie') {
      return baseOptions;
    }

    return {
      ...baseOptions,
      scales: {
        x: {
          ticks: { color: darkMode ? '#ffffff' : '#0a2540' },
          grid: { color: darkMode ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.05)' }
        },
        y: {
          ticks: { color: darkMode ? '#ffffff' : '#0a2540' },
          grid: { color: darkMode ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.05)' }
        },
        ...(selectedY.length === 2 && {
          y1: {
            position: 'right',
            ticks: { color: darkMode ? '#ffffff' : '#0a2540' },
            grid: { display: false }
          }
        })
      }
    };
  };

  const filteredData = useMemo(() => {
    if (!data) return data;

    let result = data;

    if (searchTerm) {
      result = result.filter(row => {
        if (filterColumn === 'all') {
          return Object.values(row).some(val =>
            String(val).toLowerCase().includes(searchTerm.toLowerCase())
          );
        } else {
          return String(row[filterColumn]).toLowerCase().includes(searchTerm.toLowerCase());
        }
      });
    }

    if (sortConfig.key) {
      result = [...result].sort((a, b) => {
        const aVal = a[sortConfig.key];
        const bVal = b[sortConfig.key];
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);

        if (!isNaN(aNum) && !isNaN(bNum)) {
          return sortConfig.direction === 'asc' ? aNum - bNum : bNum - aNum;
        }

        const aStr = String(aVal).toLowerCase();
        const bStr = String(bVal).toLowerCase();

        if (sortConfig.direction === 'asc') {
          return aStr < bStr ? -1 : aStr > bStr ? 1 : 0;
        } else {
          return aStr > bStr ? -1 : aStr < bStr ? 1 : 0;
        }
      });
    }

    return result;
  }, [data, searchTerm, filterColumn, sortConfig]);

  const paginatedData = useMemo(() => {
    if (!filteredData) return [];
    const start = (currentPage - 1) * pageSize;
    return filteredData.slice(start, start + pageSize);
  }, [filteredData, currentPage, pageSize]);

  const totalPages = Math.ceil((filteredData?.length || 0) / pageSize);

  const handleSort = (key) => {
    setSortConfig(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  const handleYAxisToggle = (col) => {
    setSelectedY(prev =>
      prev.includes(col) 
        ? prev.filter(c => c !== col) 
        : [...prev, col].slice(0, 2)
    );
  };

  return (
    <div className={darkMode ? 'dark' : ''}>
      <div className="app-container">
        <header className="app-header">
          <div className="header-content">
            <h1>数据可视化分析</h1>
            <button 
              onClick={() => setDarkMode(!darkMode)} 
              className="theme-toggle"
              aria-label="切换主题"
            >
              {darkMode ? '☀️' : '🌙'}
            </button>
          </div>
        </header>

        <main className="main-content">
          {!data ? (
            <div className="upload-section">
              <div
                className={`upload-area ${isDragging ? 'dragging' : ''}`}
                onClick={() => fileInputRef.current?.click()}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
              >
                <div className="upload-icon">📊</div>
                <h2>上传数据文件</h2>
                <p>支持 CSV、Excel (xlsx/xls) 格式</p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  onChange={(e) => handleFileUpload(e.target.files[0])}
                />
              </div>
            </div>
          ) : (
            <>
              <div className="toolbar">
                <div className="toolbar-section">
                  <span className="file-name">📁 {fileName}</span>
                </div>
                <div className="toolbar-section">
                  <button onClick={() => {
                    setData(null);
                    setFileName('');
                    setSelectedXValues([]);
                  }} className="btn-secondary">
                    重新上传
                  </button>
                </div>
              </div>

              {stats && (
                <div className="stats-grid">
                  <div className="stat-card">
                    <div className="stat-label">总行数</div>
                    <div className="stat-value">{data.length}</div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-label">总列数</div>
                    <div className="stat-value">{headers.length}</div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-label">平均值</div>
                    <div className="stat-value">{stats.mean}</div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-label">中位数</div>
                    <div className="stat-value">{stats.median}</div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-label">最大值</div>
                    <div className="stat-value">{stats.max}</div>
                  </div>
                  <div className="stat-card">
                    <div className="stat-label">最小值</div>
                    <div className="stat-value">{stats.min}</div>
                  </div>
                </div>
              )}

              <div className="controls-panel">
                <div className="control-group">
                  <label>图表类型</label>
                  <select value={chartType} onChange={(e) => setChartType(e.target.value)}>
                    <option value="line">折线图</option>
                    <option value="bar">柱状图</option>
                    <option value="pie">饼图</option>
                  </select>
                </div>

                <div className="control-group">
                  <label>X 轴</label>
                  <select value={selectedX} onChange={(e) => setSelectedX(e.target.value)}>
                    {headers.map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>

                {xAxisUniqueValues.length > 0 && (
                  <div className="x-filter-panel">
                    <div className="x-filter-header">
                      <label>筛选 {selectedX} 的值（不选则显示全部）</label>
                      <div className="x-filter-actions">
                        <button onClick={handleSelectAllX} className="btn-filter-action">全选</button>
                        <button onClick={handleClearAllX} className="btn-filter-action">清空</button>
                      </div>
                    </div>
                    <div className="x-filter-values">
                      {xAxisUniqueValues.map(value => (
                        <label key={value} className="checkbox-label">
                          <input
                            type="checkbox"
                            checked={selectedXValues.includes(value)}
                            onChange={() => handleXValueToggle(value)}
                          />
                          <span>{value}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                )}

                <div className="control-group">
                  <label>图表标题</label>
                  <input
                    type="text"
                    value={chartTitle}
                    onChange={(e) => setChartTitle(e.target.value)}
                    placeholder="输入标题（可选）"
                  />
                </div>
              </div>

              <div className="y-axis-panel">
                <label>Y 轴（最多选择 2 个）</label>
                <div className="checkbox-grid">
                  {headers.map(h => (
                    <label key={h} className="checkbox-label">
                      <input
                        type="checkbox"
                        checked={selectedY.includes(h)}
                        onChange={() => handleYAxisToggle(h)}
                        disabled={!selectedY.includes(h) && selectedY.length >= 2}
                      />
                      <span>{h}</span>
                    </label>
                  ))}
                </div>
              </div>

              <div className="chart-section">
                {chartData && (
                  <>
                    {chartType === 'line' && <Line ref={chartRef} data={chartData} options={getChartOptions('line')} />}
                    {chartType === 'bar' && <Bar ref={chartRef} data={chartData} options={getChartOptions('bar')} />}
                    {chartType === 'pie' && selectedY.length === 1 && <Pie ref={chartRef} data={chartData} options={getChartOptions('pie')} />}
                  </>
                )}
              </div>

              <div className="filter-panel">
                <div className="filter-group">
                  <label>筛选列</label>
                  <select value={filterColumn} onChange={(e) => setFilterColumn(e.target.value)}>
                    <option value="all">全部列</option>
                    {headers.map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>

                <div className="filter-group search-group">
                  <label>搜索</label>
                  <input
                    type="text"
                    value={searchTerm}
                    onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                    placeholder="输入关键词..."
                  />
                </div>

                {searchTerm && (
                  <button 
                    onClick={() => { 
                      setSearchTerm(''); 
                      setFilterColumn('all'); 
                    }} 
                    className="btn-clear"
                  >
                    清除
                  </button>
                )}
              </div>

              <div className="table-section">
                <div className="table-header">
                  <h3>
                    数据预览
                    {searchTerm && ` (筛选结果 ${filteredData.length} 条)`}
                  </h3>
                  <div className="page-size-control">
                    <label>
                      每页显示
                      <select 
                        value={pageSize} 
                        onChange={(e) => { 
                          setPageSize(Number(e.target.value)); 
                          setCurrentPage(1); 
                        }}
                      >
                        <option value="10">10</option>
                        <option value="20">20</option>
                        <option value="50">50</option>
                        <option value="100">100</option>
                      </select>
                    </label>
                  </div>
                </div>

                <div className="table-wrapper">
                  <table>
                    <thead>
                      <tr>
                        {headers.map(h => (
                          <th key={h} onClick={() => handleSort(h)}>
                            {h}
                            {sortConfig.key === h && (
                              <span className="sort-icon">
                                {sortConfig.direction === 'asc' ? ' ↑' : ' ↓'}
                              </span>
                            )}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedData.map((row, i) => (
                        <tr key={i}>
                          {headers.map(h => <td key={h}>{row[h]}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {totalPages > 1 && (
                  <div className="pagination">
                    <button 
                      onClick={() => setCurrentPage(1)} 
                      disabled={currentPage === 1}
                    >
                      «
                    </button>
                    <button 
                      onClick={() => setCurrentPage(p => Math.max(1, p - 1))} 
                      disabled={currentPage === 1}
                    >
                      ‹
                    </button>
                    <span>第 {currentPage} / {totalPages} 页</span>
                    <button 
                      onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} 
                      disabled={currentPage === totalPages}
                    >
                      ›
                    </button>
                    <button 
                      onClick={() => setCurrentPage(totalPages)} 
                      disabled={currentPage === totalPages}
                    >
                      »
                    </button>
                  </div>
                )}

                {filteredData.length === 0 && searchTerm && (
                  <div className="no-results">未找到匹配的数据</div>
                )}
              </div>
            </>
          )}
        </main>
      </div>
    </div>
  );
}
