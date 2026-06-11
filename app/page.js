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
  const [showXFilterModal, setShowXFilterModal] = useState(false);
  const [xFilterEnabled, setXFilterEnabled] = useState(false);
  const [xFilterSearch, setXFilterSearch] = useState('');

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
    setXFilterEnabled(false);
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

  const xAxisFilterOptions = useMemo(() => {
    if (!data || !selectedX) return [];
    return data.map((row, index) => ({
      id: String(index),
      value: String(row[selectedX] ?? ''),
      rowNumber: index + 1
    }));
  }, [data, selectedX]);

  const filteredXValues = useMemo(() => {
    if (!xFilterSearch) return xAxisFilterOptions;
    return xAxisFilterOptions.filter(({ value }) => 
      value.toLowerCase().includes(xFilterSearch.toLowerCase())
    );
  }, [xAxisFilterOptions, xFilterSearch]);

  const handleXValueToggle = (id) => {
    setSelectedXValues(prev =>
      prev.includes(id)
        ? prev.filter(v => v !== id)
        : [...prev, id]
    );
  };

  const handleSelectAllX = () => {
    setSelectedXValues(filteredXValues.map(({ id }) => id));
  };

  const handleClearAllX = () => {
    setSelectedXValues([]);
  };

  const handleConfirmXFilter = () => {
    setShowXFilterModal(false);
    setXFilterSearch('');
  };  const chartData = useMemo(() => {
    if (!data || selectedY.length === 0) return null;

    const filteredByX = (xFilterEnabled && selectedXValues.length > 0)
      ? data.filter((_, index) => selectedXValues.includes(String(index)))
      : data;

    const labels = filteredByX.map(row => String(row[selectedX] ?? ''));

    if (chartType === 'pie') {
      const yKey = selectedY[0];
      const values = filteredByX.map(row => {
        const val = parseFloat(row[yKey]);
        return isNaN(val) ? 0 : val;
      });

      return {
        labels,
        datasets: [{
          data: values,
          backgroundColor: customColors,
          borderWidth: 2,
          borderColor: darkMode ? '#1f2937' : '#fff'
        }]
      };
    }

    const datasets = selectedY.map((yKey, idx) => {
      const values = filteredByX.map(row => {
        const val = parseFloat(row[yKey]);
        return isNaN(val) ? 0 : val;
      });

      return {
        label: yKey,
        data: values,
        backgroundColor: customColors[idx % customColors.length],
        borderColor: customColors[idx % customColors.length],
        borderWidth: 2,
        tension: chartType === 'line' ? 0.4 : 0
      };
    });

    return { labels, datasets };
  }, [data, selectedX, selectedY, chartType, customColors, darkMode, xFilterEnabled, selectedXValues]);

  const chartMinWidth = useMemo(() => {
    if (!chartData || chartType === 'pie') return '100%';
    const widthPerLabel = chartType === 'bar' ? 52 : 44;
    return `${Math.max(900, chartData.labels.length * widthPerLabel)}px`;
  }, [chartData, chartType]);

  const chartLegendItems = useMemo(() => {
    if (!chartData || chartType === 'pie') return [];
    return chartData.datasets.map(dataset => ({
      label: dataset.label,
      color: dataset.borderColor || dataset.backgroundColor
    }));
  }, [chartData, chartType]);

  const getChartOptions = (type) => {
    const baseOptions = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: type === 'pie',
          position: 'top',
          labels: {
            color: darkMode ? '#e5e7eb' : '#374151',
            font: { size: 13 }
          }
        },
        title: {
          display: !!chartTitle,
          text: chartTitle,
          color: darkMode ? '#f3f4f6' : '#111827',
          font: { size: 18, weight: 'bold' }
        }
      }
    };

    if (type === 'pie') return baseOptions;

    return {
      ...baseOptions,
      scales: {
        x: {
          ticks: { 
            color: darkMode ? '#d1d5db' : '#4b5563',
            font: { size: 12 },
            autoSkip: false,
            maxRotation: 45,
            minRotation: 0
          },
          grid: { 
            color: darkMode ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)' 
          }
        },
        y: {
          ticks: { 
            color: darkMode ? '#d1d5db' : '#4b5563',
            font: { size: 12 }
          },
          grid: { 
            color: darkMode ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)' 
          }
        }
      }
    };
  };

  const filteredData = useMemo(() => {
    if (!data) return [];
    if (!searchTerm) return data;

    return data.filter(row => {
      if (filterColumn === 'all') {
        return Object.values(row).some(val =>
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        );
      }
      return String(row[filterColumn]).toLowerCase().includes(searchTerm.toLowerCase());
    });
  }, [data, searchTerm, filterColumn]);

  const sortedData = useMemo(() => {
    if (!sortConfig.key) return filteredData;

    return [...filteredData].sort((a, b) => {
      const aVal = a[sortConfig.key];
      const bVal = b[sortConfig.key];
      const aNum = parseFloat(aVal);
      const bNum = parseFloat(bVal);

      let comparison = 0;
      if (!isNaN(aNum) && !isNaN(bNum)) {
        comparison = aNum - bNum;
      } else {
        comparison = String(aVal).localeCompare(String(bVal));
      }

      return sortConfig.direction === 'asc' ? comparison : -comparison;
    });
  }, [filteredData, sortConfig]);

  const totalPages = Math.ceil(sortedData.length / pageSize);
  const paginatedData = sortedData.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );

  const handleSort = (key) => {
    setSortConfig(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  const handleYToggle = (header) => {
    setSelectedY(prev =>
      prev.includes(header)
        ? prev.filter(h => h !== header)
        : [...prev, header]
    );
  };  return (
    <div className={`app-container ${darkMode ? 'dark' : ''}`}>
      <div className="main-content">
        <header className="app-header">
          <div className="header-left">
            <h1>数据可视化分析平台</h1>
            {fileName && <span className="file-badge">{fileName}</span>}
          </div>
          <button 
            className="theme-toggle" 
            onClick={() => setDarkMode(!darkMode)}
            title="切换主题"
          >
            {darkMode ? '☀️' : '🌙'}
          </button>
        </header>

        <main className="app-main">
          {!data ? (
            <div 
              className={`upload-zone ${isDragging ? 'dragging' : ''}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <div className="upload-icon">📊</div>
              <h2>上传数据文件</h2>
              <p>拖拽文件到此处，或点击选择文件</p>
              <p className="file-hint">支持 .xlsx, .xls, .csv 格式</p>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => handleFileUpload(e.target.files[0])}
                style={{ display: 'none' }}
              />
            </div>
          ) : (
            <>
              <div className="controls-section">
                <div className="control-group">
                  <label>图表类型</label>
                  <div className="button-group">
                    <button
                      className={chartType === 'line' ? 'active' : ''}
                      onClick={() => setChartType('line')}
                    >
                      📈 折线图
                    </button>
                    <button
                      className={chartType === 'bar' ? 'active' : ''}
                      onClick={() => setChartType('bar')}
                    >
                      📊 柱状图
                    </button>
                    <button
                      className={chartType === 'pie' ? 'active' : ''}
                      onClick={() => setChartType('pie')}
                      disabled={selectedY.length > 1}
                    >
                      🥧 饼图
                    </button>
                  </div>
                </div>

                <div className="control-group">
                  <label>图表标题（可选）</label>
                  <input
                    type="text"
                    value={chartTitle}
                    onChange={(e) => setChartTitle(e.target.value)}
                    placeholder="输入图表标题..."
                  />
                </div>

                <div className="axis-controls">
                  <div className="control-group">
                    <label>X轴选择</label>
                    <select value={selectedX} onChange={(e) => {
                      setSelectedX(e.target.value);
                      setSelectedXValues([]);
                      setXFilterEnabled(false);
                    }}>
                      {headers.map(h => <option key={h} value={h}>{h}</option>)}
                    </select>
                  </div>

                  <div className="control-group x-filter-group">
                    <label>X轴筛选</label>
                    <div className={`x-filter-control ${xFilterEnabled ? 'active' : ''}`}>
                      <label className="filter-toggle">
                        <input
                          type="checkbox"
                          checked={xFilterEnabled}
                          onChange={(e) => {
                            setXFilterEnabled(e.target.checked);
                            if (!e.target.checked) {
                              setSelectedXValues([]);
                            }
                          }}
                        />
                        <span>启用</span>
                      </label>
                      <button
                        className="btn-filter"
                        disabled={!xFilterEnabled}
                        onClick={() => setShowXFilterModal(true)}
                      >
                        {xFilterEnabled
                          ? `筛选数据 (${selectedXValues.length}/${xAxisFilterOptions.length})`
                          : '先启用筛选'}
                      </button>
                    </div>
                  </div>

                  <div className="control-group">
                    <label>Y轴选择（可多选）</label>
                    <div className="y-axis-dropdown">
                      <div className="dropdown-trigger">
                        <div className="selected-values">
                          {selectedY.length === 0 ? '请选择Y轴' : selectedY.join(', ')}
                        </div>
                        <span className="dropdown-arrow">▼</span>
                      </div>
                      <div className="dropdown-menu">
                        {headers.map(h => (
                          <label key={h} className="dropdown-item">
                            <input
                              type="checkbox"
                              checked={selectedY.includes(h)}
                              onChange={() => handleYToggle(h)}
                            />
                            <span>{h}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>

                <button 
                  onClick={() => {
                    setData(null);
                    setFileName('');
                    setHeaders([]);
                    setSelectedX('');
                    setSelectedY([]);
                    setChartTitle('');
                    setSelectedXValues([]);
                    setXFilterEnabled(false);
                  }} 
                  className="btn-reset"
                >
                  重新上传
                </button>
              </div>

              {showXFilterModal && (
                <div className="modal-overlay" onClick={() => setShowXFilterModal(false)}>
                  <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                    <div className="modal-header">
                      <h3>筛选 {selectedX} 的值</h3>
                      <button className="modal-close" onClick={() => setShowXFilterModal(false)}>×</button>
                    </div>
                    <div className="modal-body">
                      <div className="filter-search">
                        <input
                          type="text"
                          placeholder="输入关键词搜索..."
                          value={xFilterSearch}
                          onChange={(e) => setXFilterSearch(e.target.value)}
                        />
                      </div>
                      <div className="filter-actions">
                        <button onClick={handleSelectAllX}>全选</button>
                        <button onClick={handleClearAllX}>清空</button>
                        <span className="selection-count">
                          已选: {selectedXValues.length} / {filteredXValues.length}
                        </span>
                      </div>
                      <div className="filter-list">
                        {filteredXValues.map(({ id, value, rowNumber }) => (
                          <label key={id} className="filter-item">
                            <input
                              type="checkbox"
                              checked={selectedXValues.includes(id)}
                              onChange={() => handleXValueToggle(id)}
                            />
                            <span className="filter-item-value">{value}</span>
                            <span className="filter-item-row">第 {rowNumber} 行</span>
                          </label>
                        ))}
                      </div>
                    </div>
                    <div className="modal-footer">
                      <button className="btn-confirm" onClick={handleConfirmXFilter}>确定</button>
                    </div>
                  </div>
                </div>
              )}              <div className="chart-section">
                {chartLegendItems.length > 0 && (
                  <div className="chart-legend" aria-label="图例">
                    {chartLegendItems.map(item => (
                      <div className="chart-legend-item" key={item.label}>
                        <span
                          className="chart-legend-swatch"
                          style={{ backgroundColor: item.color }}
                        />
                        <span>{item.label}</span>
                      </div>
                    ))}
                  </div>
                )}
                <div className="chart-scroll">
                  <div className="chart-container" style={{ minWidth: chartMinWidth }}>
                  {!chartData ? (
                    <div className="chart-placeholder">
                      请选择X轴和Y轴数据
                    </div>
                  ) : (
                    <>
                      {chartType === 'line' && <Line ref={chartRef} data={chartData} options={getChartOptions('line')} />}
                      {chartType === 'bar' && <Bar ref={chartRef} data={chartData} options={getChartOptions('bar')} />}
                      {chartType === 'pie' && selectedY.length === 1 && <Pie ref={chartRef} data={chartData} options={getChartOptions('pie')} />}
                    </>
                  )}
                  </div>
                </div>
              </div>

              <div className="filter-panel">
                <div className="filter-group filter-column-group">
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
                    {searchTerm && ` (筛选结果: ${filteredData.length} 条)`}
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
