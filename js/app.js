// 全局变量
let excelData = null;
let headers = [];
let currentPage = 1;
let rowsPerPage = 10;
let filteredData = [];
let barChart = null;
let pieChart = null;

// 页面加载完成后执行
document.addEventListener('DOMContentLoaded', function() {
    // 初始化上传功能
    initUpload();
    
    // 绑定事件监听器
    document.getElementById('searchBtn').addEventListener('click', searchData);
    document.getElementById('applyFilters').addEventListener('click', applyFilters);
    document.getElementById('resetFilters').addEventListener('click', resetFilters);
    document.getElementById('prevPage').addEventListener('click', goToPrevPage);
    document.getElementById('nextPage').addEventListener('click', goToNextPage);
    document.getElementById('exportCSV').addEventListener('click', exportToCSV);
    document.getElementById('exportExcel').addEventListener('click', exportToExcel);
    document.getElementById('printData').addEventListener('click', printData);
    
    // 加载示例数据
    loadSampleData();
});

// 初始化文件上传功能
function initUpload() {
    // 创建文件上传元素
    const uploadContainer = document.createElement('div');
    uploadContainer.className = 'card mt-3';
    uploadContainer.innerHTML = `
        <div class="card-header">
            <h5 class="card-title mb-0">上传Excel文件</h5>
        </div>
        <div class="card-body">
            <div class="mb-3">
                <input class="form-control" type="file" id="fileUpload" accept=".xlsx, .xls, .csv">
            </div>
            <div class="d-grid">
                <button class="btn btn-success" id="uploadBtn">
                    <i class="bi bi-upload me-1"></i>上传文件
                </button>
            </div>
        </div>
    `;
    
    // 插入到左侧面板的顶部
    const leftPanel = document.querySelector('.col-md-3');
    leftPanel.insertBefore(uploadContainer, leftPanel.firstChild);
    
    // 绑定上传事件
    document.getElementById('uploadBtn').addEventListener('click', handleFileUpload);
}

// 处理文件上传
function handleFileUpload() {
    const fileInput = document.getElementById('fileUpload');
    const file = fileInput.files[0];
    
    if (!file) {
        alert('请先选择一个Excel文件');
        return;
    }
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 获取第一个工作表
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // 转换为JSON
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // 处理数据
            processData(jsonData);
            
        } catch (error) {
            console.error('解析Excel文件时出错:', error);
            alert('无法解析文件，请确保上传了有效的Excel文件');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// 加载示例数据
function loadSampleData() {
    // 示例数据
    const sampleData = [
        ['产品ID', '产品名称', '类别', '价格', '库存', '销量', '评分'],
        ['P001', '笔记本电脑', '电子产品', 5999, 120, 78, 4.7],
        ['P002', '智能手机', '电子产品', 3999, 200, 156, 4.5],
        ['P003', '无线耳机', '配件', 899, 300, 210, 4.8],
        ['P004', '机械键盘', '配件', 499, 150, 89, 4.6],
        ['P005', '显示器', '电子产品', 1299, 80, 45, 4.4],
        ['P006', '游戏鼠标', '配件', 299, 200, 134, 4.7],
        ['P007', '平板电脑', '电子产品', 2999, 100, 67, 4.5],
        ['P008', '移动电源', '配件', 199, 400, 278, 4.3],
        ['P009', '智能手表', '电子产品', 1599, 120, 98, 4.6],
        ['P010', '蓝牙音箱', '电子产品', 699, 180, 123, 4.4],
        ['P011', '摄像头', '配件', 399, 150, 87, 4.2],
        ['P012', '路由器', '电子产品', 299, 200, 145, 4.5],
        ['P013', '固态硬盘', '配件', 599, 120, 76, 4.8],
        ['P014', '游戏手柄', '配件', 349, 160, 110, 4.6],
        ['P015', '智能音箱', '电子产品', 899, 90, 56, 4.3]
    ];
    
    processData(sampleData);
}

// 处理数据
function processData(data) {
    if (!data || data.length < 2) {
        alert('数据格式不正确或为空');
        return;
    }
    
    // 提取表头和数据
    headers = data[0];
    excelData = data.slice(1);
    filteredData = [...excelData];
    
    // 更新统计信息
    updateStats();
    
    // 填充筛选和排序下拉框
    populateDropdowns();
    
    // 渲染表格
    renderTable();
    
    // 创建图表
    createCharts();
}

// 更新统计信息
function updateStats() {
    document.getElementById('totalRows').textContent = excelData.length;
    document.getElementById('totalColumns').textContent = headers.length;
    document.getElementById('filteredRows').textContent = filteredData.length;
}

// 填充下拉框
function populateDropdowns() {
    const filterColumn = document.getElementById('filterColumn');
    const sortColumn = document.getElementById('sortColumn');
    
    // 清空现有选项
    filterColumn.innerHTML = '<option selected>选择列...</option>';
    sortColumn.innerHTML = '<option selected>选择列...</option>';
    
    // 添加新选项
    headers.forEach((header, index) => {
        filterColumn.innerHTML += `<option value="${index}">${header}</option>`;
        sortColumn.innerHTML += `<option value="${index}">${header}</option>`;
    });
    
    // 监听筛选列变化
    filterColumn.addEventListener('change', updateFilterValues);
}

// 更新筛选值下拉框
function updateFilterValues() {
    const filterColumn = document.getElementById('filterColumn');
    const filterValue = document.getElementById('filterValue');
    const columnIndex = filterColumn.value;
    
    // 如果选择了"选择列..."，则清空值下拉框
    if (columnIndex === '选择列...') {
        filterValue.innerHTML = '<option selected>选择值...</option>';
        return;
    }
    
    // 获取所选列的唯一值
    const uniqueValues = [...new Set(excelData.map(row => row[columnIndex]))];
    
    // 清空现有选项
    filterValue.innerHTML = '<option selected>选择值...</option>';
    
    // 添加新选项
    uniqueValues.forEach(value => {
        filterValue.innerHTML += `<option value="${value}">${value}</option>`;
    });
}

// 渲染表格
function renderTable() {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');
    
    // 渲染表头
    tableHeader.innerHTML = '';
    headers.forEach(header => {
        tableHeader.innerHTML += `<th>${header}</th>`;
    });
    
    // 计算分页
    const startIndex = (currentPage - 1) * rowsPerPage;
    const endIndex = Math.min(startIndex + rowsPerPage, filteredData.length);
    const pageData = filteredData.slice(startIndex, endIndex);
    
    // 渲染表格内容
    tableBody.innerHTML = '';
    pageData.forEach(row => {
        let tr = '<tr>';
        row.forEach(cell => {
            tr += `<td>${cell}</td>`;
        });
        tr += '</tr>';
        tableBody.innerHTML += tr;
    });
    
    // 更新分页信息
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    document.getElementById('pageInfo').textContent = `第${currentPage}页/共${totalPages}页`;
    
    // 更新分页按钮状态
    document.getElementById('prevPage').disabled = currentPage === 1;
    document.getElementById('nextPage').disabled = currentPage === totalPages;
}

// 创建图表
function createCharts() {
    createBarChart();
    createPieChart();
}

// 创建柱状图
function createBarChart() {
    // 销量数据
    const labels = filteredData.slice(0, 5).map(row => row[1]); // 产品名称
    const data = filteredData.slice(0, 5).map(row => row[5]);   // 销量
    
    const ctx = document.getElementById('barChart').getContext('2d');
    
    // 如果图表已存在，销毁它
    if (barChart) {
        barChart.destroy();
    }
    
    // 创建新图表
    barChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '销量',
                data: data,
                backgroundColor: 'rgba(58, 123, 213, 0.7)',
                borderColor: 'rgba(58, 123, 213, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: '前5个产品销量'
                }
            }
        }
    });
}

// 创建饼图
function createPieChart() {
    // 按类别统计产品数量
    const categories = {};
    filteredData.forEach(row => {
        const category = row[2]; // 类别
        categories[category] = (categories[category] || 0) + 1;
    });
    
    const labels = Object.keys(categories);
    const data = Object.values(categories);
    
    const ctx = document.getElementById('pieChart').getContext('2d');
    
    // 如果图表已存在，销毁它
    if (pieChart) {
        pieChart.destroy();
    }
    
    // 创建新图表
    pieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: [
                    'rgba(58, 123, 213, 0.7)',
                    'rgba(54, 162, 235, 0.7)',
                    'rgba(75, 192, 192, 0.7)',
                    'rgba(153, 102, 255, 0.7)',
                    'rgba(255, 159, 64, 0.7)'
                ],
                borderColor: [
                    'rgba(58, 123, 213, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: '产品类别分布'
                }
            }
        }
    });
}

// 搜索数据
function searchData() {
    const searchInput = document.getElementById('searchInput').value.toLowerCase();
    
    if (!searchInput) {
        filteredData = [...excelData];
    } else {
        filteredData = excelData.filter(row => {
            return row.some(cell => 
                String(cell).toLowerCase().includes(searchInput)
            );
        });
    }
    
    // 重置分页
    currentPage = 1;
    
    // 更新统计信息
    updateStats();
    
    // 重新渲染表格
    renderTable();
    
    // 更新图表
    createCharts();
}

// 应用筛选
function applyFilters() {
    const filterColumn = document.getElementById('filterColumn');
    const filterValue = document.getElementById('filterValue');
    const sortColumn = document.getElementById('sortColumn');
    const sortOrder = document.getElementById('sortOrder');
    
    // 重置筛选数据
    filteredData = [...excelData];
    
    // 应用列筛选
    if (filterColumn.value !== '选择列...' && filterValue.value !== '选择值...') {
        const columnIndex = parseInt(filterColumn.value);
        const value = filterValue.value;
        
        filteredData = filteredData.filter(row => 
            String(row[columnIndex]) === String(value)
        );
    }
    
    // 应用排序
    if (sortColumn.value !== '选择列...') {
        const columnIndex = parseInt(sortColumn.value);
        const order = sortOrder.value;
        
        filteredData.sort((a, b) => {
            const valueA = a[columnIndex];
            const valueB = b[columnIndex];
            
            // 数字排序
            if (!isNaN(valueA) && !isNaN(valueB)) {
                return order === 'asc' ? valueA - valueB : valueB - valueA;
            }
            
            // 字符串排序
            return order === 'asc' 
                ? String(valueA).localeCompare(String(valueB))
                : String(valueB).localeCompare(String(valueA));
        });
    }
    
    // 重置分页
    currentPage = 1;
    
    // 更新统计信息
    updateStats();
    
    // 重新渲染表格
    renderTable();
    
    // 更新图表
    createCharts();
}

// 重置筛选
function resetFilters() {
    // 重置筛选控件
    document.getElementById('searchInput').value = '';
    document.getElementById('filterColumn').selectedIndex = 0;
    document.getElementById('filterValue').innerHTML = '<option selected>选择值...</option>';
    document.getElementById('sortColumn').selectedIndex = 0;
    document.getElementById('sortOrder').selectedIndex = 0;
    
    // 重置筛选数据
    filteredData = [...excelData];
    
    // 重置分页
    currentPage = 1;
    
    // 更新统计信息
    updateStats();
    
    // 重新渲染表格
    renderTable();
    
    // 更新图表
    createCharts();
}

// 上一页
function goToPrevPage() {
    if (currentPage > 1) {
        currentPage--;
        renderTable();
    }
}

// 下一页
function goToNextPage() {
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    if (currentPage < totalPages) {
        currentPage++;
        renderTable();
    }
}

// 导出为CSV
function exportToCSV() {
    // 创建CSV内容
    let csvContent = headers.join(',') + '\n';
    
    filteredData.forEach(row => {
        csvContent += row.join(',') + '\n';
    });
    
    // 创建Blob对象
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    
    // 创建下载链接
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', 'exported_data.csv');
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 导出为Excel
function exportToExcel() {
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 创建工作表数据
    const wsData = [headers, ...filteredData];
    
    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    // 导出工作簿
    XLSX.writeFile(wb, 'exported_data.xlsx');
}

// 打印数据
function printData() {
    // 创建打印窗口
    const printWindow = window.open('', '_blank');
    
    // 创建打印内容
    let printContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>打印数据</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }
                th {
                    background-color: #f2f2f2;
                }
                h1 {
                    text-align: center;
                    margin-bottom: 20px;
                }
                .print-info {
                    text-align: right;
                    font-size: 12px;
                    color: #666;
                    margin-bottom: 20px;
                }
                @media print {
                    .no-print {
                        display: none;
                    }
                }
            </style>
        </head>
        <body>
            <h1>Excel数据可视化平台 - 数据报表</h1>
            <div class="print-info">
                生成时间: ${new Date().toLocaleString()}
            </div>
            <div class="no-print" style="text-align: center; margin-bottom: 20px;">
                <button onclick="window.print()">打印</button>
                <button onclick="window.close()">关闭</button>
            </div>
            <table>
                <thead>
                    <tr>
                        ${headers.map(header => `<th>${header}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>
                    ${filteredData.map(row => `
                        <tr>
                            ${row.map(cell => `<td>${cell}</td>`).join('')}
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </body>
        </html>
    `;
    
    // 写入打印内容
    printWindow.document.open();
    printWindow.document.write(printContent);
    printWindow.document.close();
}

// 辅助函数：格式化数字
function formatNumber(num) {
    return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
}

// 辅助函数：获取随机颜色
function getRandomColor() {
    const letters = '0123456789ABCDEF';
    let color = '#';
    for (let i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
}

// 辅助函数：防抖
function debounce(func, wait) {
    let timeout;
    return function(...args) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), wait);
    };
}

// 为搜索框添加防抖
const searchInput = document.getElementById('searchInput');
searchInput.addEventListener('input', debounce(function() {
    document.getElementById('searchBtn').click();
}, 300));

// 响应式调整
window.addEventListener('resize', function() {
    if (barChart) barChart.resize();
    if (pieChart) pieChart.resize();
});

// 添加表格排序功能
function addTableSorting() {
    const tableHeader = document.getElementById('tableHeader');
    const headerCells = tableHeader.querySelectorAll('th');
    
    headerCells.forEach((cell, index) => {
        cell.style.cursor = 'pointer';
        cell.addEventListener('click', () => {
            // 设置排序列和顺序
            document.getElementById('sortColumn').value = index;
            
            // 切换排序顺序
            const currentOrder = document.getElementById('sortOrder').value;
            document.getElementById('sortOrder').value = currentOrder === 'asc' ? 'desc' : 'asc';
            
            // 应用筛选
            applyFilters();
        });
    });
}

// 初始化后添加表格排序功能
setTimeout(addTableSorting, 1000);