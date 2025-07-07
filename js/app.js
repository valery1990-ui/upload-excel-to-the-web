// 全局变量
let excelData = null;
let headers = [];
let currentPage = 1;
let rowsPerPage = 10;
let filteredData = [];
let barChart = null;
let pieChart = null;

// ===== 国际化支持 =====
const i18nTexts = {
    zh: {
        exportCSV: '导出CSV',
        exportExcel: '导出Excel',
        printData: '打印',
        dataControl: '数据控制',
        searchData: '搜索数据',
        searchPlaceholder: '输入关键词...',
        filterCondition: '筛选条件',
        selectColumn: '选择列...',
        selectValue: '选择值...',
        sortMethod: '排序方式',
        asc: '升序',
        desc: '降序',
        applyFilter: '应用筛选',
        resetFilter: '重置筛选',
        dataStats: '数据统计',
        totalRows: '总行数',
        totalColumns: '总列数',
        filteredRows: '筛选后行数',
        dataTable: '数据表格',
        pageInfo: '第1页/共1页',
        dataVisualization: '数据可视化',
        footer: 'Excel数据可视化平台 © 2025',
        brandTitle: 'Excel数据可视化平台',
        // 上传区
        uploadTitle: '上传Excel文件',
        chooseFile: '选择文件',
        noFile: '未选择文件',
        uploadBtn: '上传文件',
        // 表头
        header_productId: '产品ID',
        header_productName: '产品名称',
        header_category: '类别',
        header_price: '价格',
        header_stock: '库存',
        header_sales: '销量',
        header_rating: '评分',
        // 图表
        chart_bar_title: '前5个产品销量',
        chart_bar_label: '销量',
        chart_pie_title: '产品类别分布',
        chart_pie_label1: '电子产品',
        chart_pie_label2: '配件',
        // 分页
        pageInfoFormat: '第{current}页/共{total}页',
        // 弹窗
        alert_no_file: '请先选择一个Excel文件',
        alert_parse_error: '无法解析文件，请确保上传了有效的Excel文件',
        alert_data_error: '数据格式不正确或为空',
    },
    en: {
        exportCSV: 'Export CSV',
        exportExcel: 'Export Excel',
        printData: 'Print',
        dataControl: 'Data Control',
        searchData: 'Search Data',
        searchPlaceholder: 'Enter keyword...',
        filterCondition: 'Filter Condition',
        selectColumn: 'Select column...',
        selectValue: 'Select value...',
        sortMethod: 'Sort Method',
        asc: 'Ascending',
        desc: 'Descending',
        applyFilter: 'Apply Filter',
        resetFilter: 'Reset Filter',
        dataStats: 'Statistics',
        totalRows: 'Total Rows',
        totalColumns: 'Total Columns',
        filteredRows: 'Filtered Rows',
        dataTable: 'Data Table',
        pageInfo: 'Page 1 / 1',
        dataVisualization: 'Data Visualization',
        footer: 'Excel Data Visualization Platform © 2025',
        brandTitle: 'Excel Data Visualization Platform',
        // 上传区
        uploadTitle: 'Upload Excel File',
        chooseFile: 'Choose File',
        noFile: 'No file chosen',
        uploadBtn: 'Upload File',
        // 表头
        header_productId: 'Product ID',
        header_productName: 'Product Name',
        header_category: 'Category',
        header_price: 'Price',
        header_stock: 'Stock',
        header_sales: 'Sales',
        header_rating: 'Rating',
        // 图表
        chart_bar_title: 'Top 5 Product Sales',
        chart_bar_label: 'Sales',
        chart_pie_title: 'Product Category Distribution',
        chart_pie_label1: 'Electronics',
        chart_pie_label2: 'Accessories',
        // 分页
        pageInfoFormat: 'Page {current} / {total}',
        // 弹窗
        alert_no_file: 'Please choose an Excel file first',
        alert_parse_error: 'Failed to parse file. Please upload a valid Excel file.',
        alert_data_error: 'Data format is incorrect or empty',
    }
};
let currentLang = 'en';

function setLanguage(lang) {
    currentLang = lang;
    // 切换所有 data-i18n 文本
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (i18nTexts[lang][key]) {
            el.textContent = i18nTexts[lang][key];
        }
    });
    // 切换所有 data-i18n-placeholder
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        const key = el.getAttribute('data-i18n-placeholder');
        if (i18nTexts[lang][key]) {
            el.setAttribute('placeholder', i18nTexts[lang][key]);
        }
    });
    // 切换语言按钮文本
    const langBtn = document.getElementById('langToggle');
    if (langBtn) {
        langBtn.textContent = lang === 'zh' ? 'English' : '中文';
    }
    // 刷新动态内容
    // 上传区
    const uploadTitle = document.getElementById('uploadTitle');
    if (uploadTitle) uploadTitle.textContent = i18nTexts[lang].uploadTitle;
    const uploadBtn = document.getElementById('uploadBtn');
    if (uploadBtn) uploadBtn.innerHTML = `<i class="bi bi-upload me-1"></i>${i18nTexts[lang].uploadBtn}`;
    // 文件选择按钮和文件名
    const chooseFileBtn = document.getElementById('chooseFileBtn');
    const fileNameLabel = document.getElementById('fileNameLabel');
    if (chooseFileBtn) chooseFileBtn.textContent = i18nTexts[lang].chooseFile;
    if (fileNameLabel) {
        const fileInput = document.getElementById('fileUpload');
        if (fileInput && fileInput.files.length > 0) {
            fileNameLabel.textContent = fileInput.files[0].name;
        } else {
            fileNameLabel.textContent = i18nTexts[lang].noFile;
        }
    }
    // 表格、下拉、分页、图表
    populateDropdowns();
    renderTable();
    createCharts();
}

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
    
    // 绑定语言切换按钮
    const langBtn = document.getElementById('langToggle');
    if (langBtn) {
        langBtn.addEventListener('click', function() {
            setLanguage(currentLang === 'zh' ? 'en' : 'zh');
        });
    }
    // 页面初始语言
    setLanguage(currentLang);
    
    // 加载示例数据
    loadSampleData();
});

// 初始化文件上传功能
function initUpload() {
    // 创建文件上传元素（自定义按钮和文件名显示）
    const uploadContainer = document.createElement('div');
    uploadContainer.className = 'card mt-3';
    uploadContainer.innerHTML = `
        <div class="card-header">
            <h5 class="card-title mb-0" id="uploadTitle">${i18nTexts[currentLang].uploadTitle}</h5>
        </div>
        <div class="card-body">
            <div class="mb-3">
                <div class="input-group">
                    <input type="file" class="form-control d-none" id="fileUpload" accept=".xlsx, .xls, .csv">
                    <button class="btn btn-outline-secondary" type="button" id="chooseFileBtn">${i18nTexts[currentLang].chooseFile}</button>
                    <span class="input-group-text flex-fill" id="fileNameLabel">${i18nTexts[currentLang].noFile}</span>
                </div>
            </div>
            <div class="d-grid">
                <button class="btn btn-success" id="uploadBtn">
                    <i class="bi bi-upload me-1"></i>${i18nTexts[currentLang].uploadBtn}
                </button>
            </div>
        </div>
    `;
    // 插入到左侧面板的顶部
    const leftPanel = document.querySelector('.col-md-3');
    leftPanel.insertBefore(uploadContainer, leftPanel.firstChild);
    // 绑定自定义文件选择按钮
    const fileInput = document.getElementById('fileUpload');
    const chooseFileBtn = document.getElementById('chooseFileBtn');
    const fileNameLabel = document.getElementById('fileNameLabel');
    chooseFileBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            fileNameLabel.textContent = fileInput.files[0].name;
        } else {
            fileNameLabel.textContent = i18nTexts[currentLang].noFile;
        }
    });
    // 绑定上传事件
    document.getElementById('uploadBtn').addEventListener('click', handleFileUpload);
}

// 处理文件上传
function handleFileUpload() {
    const fileInput = document.getElementById('fileUpload');
    const file = fileInput.files[0];
    
    if (!file) {
        alert(i18nTexts[currentLang].alert_no_file);
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
            alert(i18nTexts[currentLang].alert_parse_error);
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
        alert(i18nTexts[currentLang].alert_data_error);
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
    filterColumn.innerHTML = `<option selected>${i18nTexts[currentLang].selectColumn}</option>`;
    sortColumn.innerHTML = `<option selected>${i18nTexts[currentLang].selectColumn}</option>`;
    
    // 添加新选项
    headers.forEach((header, index) => {
        let key = '';
        switch (header) {
            case '产品ID': case 'Product ID': key = 'header_productId'; break;
            case '产品名称': case 'Product Name': key = 'header_productName'; break;
            case '类别': case 'Category': key = 'header_category'; break;
            case '价格': case 'Price': key = 'header_price'; break;
            case '库存': case 'Stock': key = 'header_stock'; break;
            case '销量': case 'Sales': key = 'header_sales'; break;
            case '评分': case 'Rating': key = 'header_rating'; break;
            default: key = ''; break;
        }
        filterColumn.innerHTML += `<option value="${index}">${key && i18nTexts[currentLang][key] ? i18nTexts[currentLang][key] : header}</option>`;
        sortColumn.innerHTML += `<option value="${index}">${key && i18nTexts[currentLang][key] ? i18nTexts[currentLang][key] : header}</option>`;
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
    if (columnIndex === i18nTexts[currentLang].selectColumn) {
        filterValue.innerHTML = `<option selected>${i18nTexts[currentLang].selectValue}</option>`;
        return;
    }
    
    // 获取所选列的唯一值
    const uniqueValues = [...new Set(excelData.map(row => row[columnIndex]))];
    
    // 清空现有选项
    filterValue.innerHTML = `<option selected>${i18nTexts[currentLang].selectValue}</option>`;
    
    // 添加新选项
    uniqueValues.forEach(value => {
        filterValue.innerHTML += `<option value="${value}">${value}</option>`;
    });
}

// 示例数据内容中英文映射
const productI18nMap = {
    name: {
        '笔记本电脑': 'Laptop',
        '智能手机': 'Smartphone',
        '无线耳机': 'Wireless Earbuds',
        '机械键盘': 'Mechanical Keyboard',
        '显示器': 'Monitor',
        '游戏鼠标': 'Gaming Mouse',
        '平板电脑': 'Tablet',
        '移动电源': 'Power Bank',
        '智能手表': 'Smartwatch',
        '蓝牙音箱': 'Bluetooth Speaker',
        '摄像头': 'Webcam',
        '路由器': 'Router',
        '固态硬盘': 'SSD',
        '游戏手柄': 'Gamepad',
        '智能音箱': 'Smart Speaker',
    },
    category: {
        '电子产品': 'Electronics',
        '配件': 'Accessories',
    }
};
const productI18nMapReverse = {
    name: Object.fromEntries(Object.entries(productI18nMap.name).map(([k,v])=>[v,k])),
    category: Object.fromEntries(Object.entries(productI18nMap.category).map(([k,v])=>[v,k]))
};

function translateCell(cell, colIdx, lang) {
    // 只对示例数据做映射
    if (lang === 'en') {
        if (colIdx === 1 && productI18nMap.name[cell]) return productI18nMap.name[cell];
        if (colIdx === 2 && productI18nMap.category[cell]) return productI18nMap.category[cell];
    } else if (lang === 'zh') {
        if (colIdx === 1 && productI18nMapReverse.name[cell]) return productI18nMapReverse.name[cell];
        if (colIdx === 2 && productI18nMapReverse.category[cell]) return productI18nMapReverse.category[cell];
    }
    return cell;
}

function renderTable() {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');
    // 渲染表头（国际化）
    tableHeader.innerHTML = '';
    headers.forEach((header, idx) => {
        let key = '';
        switch (header) {
            case '产品ID': key = 'header_productId'; break;
            case '产品名称': key = 'header_productName'; break;
            case '类别': key = 'header_category'; break;
            case '价格': key = 'header_price'; break;
            case '库存': key = 'header_stock'; break;
            case '销量': key = 'header_sales'; break;
            case '评分': key = 'header_rating'; break;
            case 'Product ID': key = 'header_productId'; break;
            case 'Product Name': key = 'header_productName'; break;
            case 'Category': key = 'header_category'; break;
            case 'Price': key = 'header_price'; break;
            case 'Stock': key = 'header_stock'; break;
            case 'Sales': key = 'header_sales'; break;
            case 'Rating': key = 'header_rating'; break;
            default: key = ''; break;
        }
        tableHeader.innerHTML += `<th>${key && i18nTexts[currentLang][key] ? i18nTexts[currentLang][key] : header}</th>`;
    });
    // 计算分页
    const startIndex = (currentPage - 1) * rowsPerPage;
    const endIndex = Math.min(startIndex + rowsPerPage, filteredData.length);
    const pageData = filteredData.slice(startIndex, endIndex);
    // 渲染表格内容（国际化）
    tableBody.innerHTML = '';
    pageData.forEach(row => {
        let tr = '<tr>';
        row.forEach((cell, colIdx) => {
            tr += `<td>${translateCell(cell, colIdx, currentLang)}</td>`;
        });
        tr += '</tr>';
        tableBody.innerHTML += tr;
    });
    // 更新分页信息（国际化）
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    let pageInfoText = i18nTexts[currentLang].pageInfoFormat
        .replace('{current}', currentPage)
        .replace('{total}', totalPages);
    document.getElementById('pageInfo').textContent = pageInfoText;
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
    // labels: 产品名称
    const labels = filteredData.slice(0, 5).map(row => translateCell(row[1], 1, currentLang));
    const data = filteredData.slice(0, 5).map(row => row[5]);
    
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
                label: i18nTexts[currentLang].chart_bar_label,
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
                    text: i18nTexts[currentLang].chart_bar_title
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
        const category = translateCell(row[2], 2, currentLang);
        categories[category] = (categories[category] || 0) + 1;
    });
    
    // 国际化类别标签
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
                    text: i18nTexts[currentLang].chart_pie_title
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
    if (filterColumn.value !== i18nTexts[currentLang].selectColumn && filterValue.value !== i18nTexts[currentLang].selectValue) {
        const columnIndex = parseInt(filterColumn.value);
        const value = filterValue.value;
        
        filteredData = filteredData.filter(row => 
            String(row[columnIndex]) === String(value)
        );
    }
    
    // 应用排序
    if (sortColumn.value !== i18nTexts[currentLang].selectColumn) {
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
    document.getElementById('filterValue').innerHTML = `<option selected>${i18nTexts[currentLang].selectValue}</option>`;
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