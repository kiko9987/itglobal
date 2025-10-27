import logger from '../utils/logger.js';

/**
 * Stats Page Manager
 * 매출 통계 페이지 관리 클래스
 */

class StatsManager {
    constructor() {
        this.projectsData = [];
        this.charts = {};
        this.managerColorMap = {};
        this.baseColors = ['#FF7F7F', '#FF9999', '#FFCC99', '#D4E6B7', '#99E6CC', '#B3C6DB', '#FF9999', '#FF6B6B'];
        this.additionalColors = ['#99D6B7', '#FFCCBB', '#D6E6E6', '#A6D1E6', '#E6D199', '#E6C299', '#E6B899', '#E6B3B3'];
        this.colorIndex = 0;
        this.additionalColorIndex = 0;
    }

    async init() {
        await this.loadStatsData();
    }

    async loadStatsData() {
        try {
            this.showLoading(true);

            const response = await fetch('/api/projects/list');
            const result = await response.json();

            // API 응답 포맷 호환성 처리
            if (Array.isArray(result)) {
                this.projectsData = result;
            } else if (result.success && Array.isArray(result.data)) {
                this.projectsData = result.data;
            } else if (result.error) {
                throw new Error(result.error);
            } else {
                throw new Error('올바르지 않은 API 응답 형태입니다');
            }

            this.generateStatistics();
            this.createCharts();
            this.showContent();
        } catch (error) {
            logger.error('통계 데이터 로드 오류:', error);
            alert('통계 데이터를 불러올 수 없습니다: ' + error.message);
        } finally {
            this.showLoading(false);
        }
    }

    generateStatistics() {
        // 2024년 프로젝트 수 계산
        const projects2024Count = this.projectsData.filter(project => {
            const confirmDate = project['공사 확정'];
            if (confirmDate) {
                const year = new Date(confirmDate).getFullYear();
                return year === 2024;
            }
            return false;
        }).length;

        // 2025년 프로젝트 수 계산
        const projects2025Count = this.projectsData.filter(project => {
            const confirmDate = project['공사 확정'];
            if (confirmDate) {
                const year = new Date(confirmDate).getFullYear();
                return year === 2025;
            }
            return false;
        }).length;

        // 총 프로젝트 표시
        document.getElementById('totalProjects').innerHTML =
            `2024년 : ${projects2024Count}건 (1월~12월)<br>2025년 : ${projects2025Count}건 (1월~현재)`;

        // 2024년 매출 계산
        const revenue2024 = this.calculateYearlyRevenue(2024);
        const revenue2025 = this.calculateYearlyRevenue(2025);

        // 총 매출 표시
        document.getElementById('totalRevenue').innerHTML =
            `2024년 : ${this.formatCurrency(revenue2024)} (1월~12월)<br>2025년 : ${this.formatCurrency(revenue2025)} (1월~현재)`;

        // 미수금 계산
        const outstanding2024 = this.calculateOutstanding(2024);
        const outstanding2025 = this.calculateOutstanding(2025);

        document.getElementById('outstandingAmount').innerHTML =
            `2024년 : ${this.formatCurrency(outstanding2024)} (1월~12월)<br>2025년 : ${this.formatCurrency(outstanding2025)} (1월~현재)`;

        // 월 평균 매출 계산
        const currentMonth = new Date().getMonth() + 1;
        const avg2024 = revenue2024 > 0 ? revenue2024 / 12 : 0;
        const avg2025 = revenue2025 > 0 ? revenue2025 / currentMonth : 0;

        document.getElementById('avgMonthlyRevenue').innerHTML =
            `2024년 : ${this.formatCurrency(avg2024)} (12개월)<br>2025년 : ${this.formatCurrency(avg2025)} (${currentMonth}개월)`;
    }

    calculateYearlyRevenue(year) {
        return this.projectsData
            .filter(project => {
                const confirmDate = project['공사 확정'];
                if (confirmDate) {
                    const projectYear = new Date(confirmDate).getFullYear();
                    return projectYear === year;
                }
                return false;
            })
            .reduce((sum, project) => {
                const amount = parseFloat(project['총액1'] || project['총액 1'] || 0);
                return sum + amount;
            }, 0);
    }

    calculateOutstanding(year) {
        return this.projectsData
            .filter(project => {
                const confirmDate = project['공사 확정'];
                if (confirmDate) {
                    const projectYear = new Date(confirmDate).getFullYear();
                    return projectYear === year;
                }
                return false;
            })
            .reduce((sum, project) => {
                const outstanding = parseFloat(project['미수금'] || project['미수금W'] || project['미수금 W'] || project['W'] || 0);
                return sum + Math.max(0, outstanding);
            }, 0);
    }

    setupManagerColors() {
        const excludedOwners = ['김단이', '심장원', '아이티', '이근혁', '황샛별'];
        const managers2024 = new Set();
        const managers2025 = new Set();

        // 담당자 수집
        this.projectsData.forEach(project => {
            const confirmDate = project['공사 확정'];
            if (confirmDate) {
                const year = new Date(confirmDate).getFullYear();
                const manager = project['담당자'];
                if (manager && !excludedOwners.includes(manager)) {
                    if (year === 2024) managers2024.add(manager);
                    if (year === 2025) managers2025.add(manager);
                }
            }
        });

        // 2024년 담당자에게 기본 색상 할당
        Array.from(managers2024).sort().forEach(manager => {
            if (this.colorIndex < this.baseColors.length) {
                this.managerColorMap[manager] = this.baseColors[this.colorIndex++];
            }
        });

        // 2025년 신규 담당자에게 추가 색상 할당
        Array.from(managers2025).sort().forEach(manager => {
            if (!this.managerColorMap[manager]) {
                if (this.additionalColorIndex < this.additionalColors.length) {
                    this.managerColorMap[manager] = this.additionalColors[this.additionalColorIndex++];
                }
            }
        });
    }

    createCharts() {
        this.setupManagerColors();
        this.createMonthly2024Chart();
        this.createMonthly2025Chart();
        this.createPersonal2024Chart();
        this.createPersonal2025Chart();
        this.createCompany2024Chart();
        this.createCompany2025Chart();
        this.createCompanyCompareChart();
        this.createTrendChart();
    }

    createMonthly2024Chart() {
        const ctx = document.getElementById('monthly2024Chart').getContext('2d');
        const monthlyData = this.getMonthlyData(2024);
        const monthNames = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];

        if (this.charts.monthly2024) this.charts.monthly2024.destroy();

        this.charts.monthly2024 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: monthNames,
                datasets: [{
                    data: monthlyData.map(amount => amount / 100000000),
                    backgroundColor: '#7BB3FF',
                    borderColor: '#5A9BD4',
                    borderWidth: 1
                }]
            },
            options: this.getBarChartOptions()
        });
    }

    createMonthly2025Chart() {
        const ctx = document.getElementById('monthly2025Chart').getContext('2d');
        const monthlyData = this.getMonthlyData(2025);
        const monthNames = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];

        if (this.charts.monthly2025) this.charts.monthly2025.destroy();

        this.charts.monthly2025 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: monthNames,
                datasets: [{
                    data: monthlyData.map(amount => amount / 100000000),
                    backgroundColor: '#7BB3FF',
                    borderColor: '#5A9BD4',
                    borderWidth: 1
                }]
            },
            options: this.getBarChartOptions()
        });
    }

    createPersonal2024Chart() {
        const ctx = document.getElementById('personal2024Chart').getContext('2d');
        const managerRevenue = this.getManagerRevenue(2024);

        const sortedManagers = Object.entries(managerRevenue)
            .sort(([, a], [, b]) => b - a);

        const managerNames = sortedManagers.map(([name]) => name);
        const managerAmounts = sortedManagers.map(([, amount]) => amount / 100000000);
        const managerColors = managerNames.map(name => this.managerColorMap[name] || '#7BB3FF');

        if (this.charts.personal2024) this.charts.personal2024.destroy();

        this.charts.personal2024 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: managerNames,
                datasets: [{
                    data: managerAmounts,
                    backgroundColor: managerColors,
                    borderColor: managerColors.map(color => color.replace('FF', 'CC')),
                    borderWidth: 1
                }]
            },
            options: this.getPersonalChartOptions()
        });
    }

    createPersonal2025Chart() {
        const ctx = document.getElementById('personal2025Chart').getContext('2d');
        const managerRevenue = this.getManagerRevenue(2025);

        const sortedManagers = Object.entries(managerRevenue)
            .sort(([, a], [, b]) => b - a);

        const managerNames = sortedManagers.map(([name]) => name);
        const managerAmounts = sortedManagers.map(([, amount]) => amount / 100000000);
        const managerColors = managerNames.map(name => this.managerColorMap[name] || '#999999');

        if (this.charts.personal2025) this.charts.personal2025.destroy();

        this.charts.personal2025 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: managerNames,
                datasets: [{
                    data: managerAmounts,
                    backgroundColor: managerColors,
                    borderColor: managerColors.map(color => color.replace('FF', 'CC')),
                    borderWidth: 1
                }]
            },
            options: this.getPersonalChartOptions()
        });
    }

    createCompany2024Chart() {
        const ctx = document.getElementById('company2024Chart').getContext('2d');
        const { globalData, globalGroupData } = this.getCompanyMonthlyData(2024);

        if (this.charts.company2024) this.charts.company2024.destroy();

        this.charts.company2024 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
                datasets: [{
                    label: '글로벌',
                    data: globalData.map(amount => amount / 100000000),
                    backgroundColor: '#FF8A95',
                    borderColor: '#E6707A',
                    borderWidth: 1
                }, {
                    label: '글로벌그룹',
                    data: globalGroupData.map(amount => amount / 100000000),
                    backgroundColor: '#7BB3FF',
                    borderColor: '#5A9BD4',
                    borderWidth: 1
                }]
            },
            options: this.getCompanyChartOptions()
        });
    }

    createCompany2025Chart() {
        const ctx = document.getElementById('company2025Chart').getContext('2d');
        const { globalData, globalGroupData } = this.getCompanyMonthlyData(2025);

        if (this.charts.company2025) this.charts.company2025.destroy();

        this.charts.company2025 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
                datasets: [{
                    label: '글로벌',
                    data: globalData.map(amount => amount / 100000000),
                    backgroundColor: '#FF8A95',
                    borderColor: '#E6707A',
                    borderWidth: 1
                }, {
                    label: '글로벌그룹',
                    data: globalGroupData.map(amount => amount / 100000000),
                    backgroundColor: '#7BB3FF',
                    borderColor: '#5A9BD4',
                    borderWidth: 1
                }]
            },
            options: this.getCompanyChartOptions()
        });
    }

    createCompanyCompareChart() {
        const ctx = document.getElementById('companyCompareChart').getContext('2d');
        const comparison = this.getCompanyComparison();

        if (this.charts.companyCompare) this.charts.companyCompare.destroy();

        this.charts.companyCompare = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['2024년', '2025년'],
                datasets: [{
                    label: '글로벌',
                    data: [comparison.global2024 / 100000000, comparison.global2025 / 100000000],
                    backgroundColor: '#FF8A95',
                    borderColor: '#E6707A',
                    borderWidth: 1
                }, {
                    label: '글로벌그룹',
                    data: [comparison.globalGroup2024 / 100000000, comparison.globalGroup2025 / 100000000],
                    backgroundColor: '#7BB3FF',
                    borderColor: '#5A9BD4',
                    borderWidth: 1
                }]
            },
            options: this.getCompanyCompareChartOptions()
        });
    }

    createTrendChart() {
        const ctx = document.getElementById('trendChart').getContext('2d');
        const monthlyData2024 = this.getMonthlyData(2024);
        const monthlyData2025 = this.getMonthlyData(2025);

        if (this.charts.trend) this.charts.trend.destroy();

        this.charts.trend = new Chart(ctx, {
            type: 'line',
            data: {
                labels: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
                datasets: [{
                    label: '2024년',
                    data: monthlyData2024.map(amount => amount / 100000000),
                    borderColor: '#FF6B35',
                    backgroundColor: 'rgba(255, 107, 53, 0.1)',
                    borderWidth: 2,
                    fill: false
                }, {
                    label: '2025년',
                    data: monthlyData2025.map(amount => amount / 100000000),
                    borderColor: '#4A90E2',
                    backgroundColor: 'rgba(74, 144, 226, 0.1)',
                    borderWidth: 2,
                    fill: false
                }]
            },
            options: this.getTrendChartOptions()
        });
    }

    // 헬퍼 메서드들
    getMonthlyData(year) {
        const monthlyData = Array(12).fill(0);

        this.projectsData.forEach(project => {
            const confirmDate = project['공사 확정'];
            if (confirmDate) {
                const confirmDateObj = new Date(confirmDate);
                if (confirmDateObj.getFullYear() === year) {
                    const month = confirmDateObj.getMonth();
                    const amount = parseFloat(project['총액1'] || project['총액 1'] || 0);
                    monthlyData[month] += amount;
                }
            }
        });

        return monthlyData;
    }

    getManagerRevenue(year) {
        const excludedOwners = ['김단이', '심장원', '아이티', '이근혁', '황샛별'];
        const managerRevenue = {};

        this.projectsData.forEach(project => {
            const confirmDate = project['공사 확정'];
            if (confirmDate) {
                const confirmDateObj = new Date(confirmDate);
                if (confirmDateObj.getFullYear() === year) {
                    const manager = project['담당자'];
                    if (manager && !excludedOwners.includes(manager)) {
                        const amount = parseFloat(project['총액1'] || project['총액 1'] || 0);
                        managerRevenue[manager] = (managerRevenue[manager] || 0) + amount;
                    }
                }
            }
        });

        return managerRevenue;
    }

    getCompanyMonthlyData(year) {
        const globalData = Array(12).fill(0);
        const globalGroupData = Array(12).fill(0);

        this.projectsData.forEach(project => {
            const confirmDate = project['공사 확정'];
            const company = project['사업자'] || '';

            if (confirmDate) {
                const confirmDateObj = new Date(confirmDate);
                if (confirmDateObj.getFullYear() === year) {
                    const month = confirmDateObj.getMonth();
                    const amount = parseFloat(project['총액1'] || project['총액 1'] || 0);

                    if (company === '글로벌') {
                        globalData[month] += amount;
                    } else if (company === '글로벌그룹') {
                        globalGroupData[month] += amount;
                    }
                }
            }
        });

        return { globalData, globalGroupData };
    }

    getCompanyComparison() {
        let global2024 = 0, globalGroup2024 = 0, global2025 = 0, globalGroup2025 = 0;

        this.projectsData.forEach(project => {
            const confirmDate = project['공사 확정'];
            const company = project['사업자'] || '';

            if (confirmDate) {
                const confirmDateObj = new Date(confirmDate);
                const year = confirmDateObj.getFullYear();
                const amount = parseFloat(project['총액1'] || project['총액 1'] || 0);

                if (company === '글로벌') {
                    if (year === 2024) global2024 += amount;
                    if (year === 2025) global2025 += amount;
                } else if (company === '글로벌그룹') {
                    if (year === 2024) globalGroup2024 += amount;
                    if (year === 2025) globalGroup2025 += amount;
                }
            }
        });

        return { global2024, globalGroup2024, global2025, globalGroup2025 };
    }

    // Chart.js 옵션들
    getBarChartOptions() {
        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: (context) => context.parsed.y.toFixed(1) + '억'
                    }
                }
            },
            scales: {
                y: {
                    min: 0,
                    max: 10,
                    ticks: {
                        stepSize: 1,
                        callback: (value) => Math.floor(value) + '억'
                    },
                    grid: {
                        color: '#e0e0e0'
                    }
                },
                x: {
                    grid: {
                        display: false
                    }
                }
            }
        };
    }

    getPersonalChartOptions() {
        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: (context) => context.label + ': ' + context.parsed.y.toFixed(1) + '억원'
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: (value) => Math.floor(value) + '억'
                    }
                },
                x: {
                    ticks: {
                        maxRotation: 0,
                        minRotation: 0
                    }
                }
            }
        };
    }

    getCompanyChartOptions() {
        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: { size: 10 }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (context) => context.dataset.label + ': ' + context.parsed.y.toFixed(1) + '억'
                    }
                }
            },
            scales: {
                y: {
                    min: 0,
                    max: 10,
                    ticks: {
                        stepSize: 1,
                        font: { size: 10 },
                        callback: (value) => Math.floor(value) + '억'
                    },
                    grid: { color: '#e0e0e0' }
                },
                x: {
                    ticks: { font: { size: 10 } },
                    grid: { display: false }
                }
            }
        };
    }

    getCompanyCompareChartOptions() {
        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: { size: 10 }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (context) => context.dataset.label + ': ' + context.parsed.y.toFixed(1) + '억'
                    }
                }
            },
            scales: {
                y: {
                    min: 0,
                    max: 100,
                    ticks: {
                        stepSize: 10,
                        font: { size: 10 },
                        callback: (value) => Math.floor(value) + '억'
                    },
                    grid: { color: '#e0e0e0' }
                },
                x: {
                    ticks: { font: { size: 10 } },
                    grid: { display: false }
                }
            }
        };
    }

    getTrendChartOptions() {
        return {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: { size: 10 }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (context) => context.dataset.label + ': ' + context.parsed.y.toFixed(1) + '억'
                    }
                }
            },
            scales: {
                y: {
                    min: 0,
                    max: 10,
                    ticks: {
                        stepSize: 1,
                        font: { size: 10 },
                        callback: (value) => Math.floor(value) + '억'
                    },
                    grid: { color: '#e0e0e0' }
                },
                x: {
                    ticks: { font: { size: 10 } },
                    grid: { display: false }
                }
            }
        };
    }

    async refresh() {
        const btn = document.getElementById('refreshStatsBtn');
        const icon = btn.querySelector('i');

        // 로딩 상태
        icon.className = 'fas fa-spinner fa-spin me-1';
        btn.disabled = true;

        try {
            // 기존 차트 삭제
            Object.values(this.charts).forEach(chart => {
                if (chart && typeof chart.destroy === 'function') {
                    chart.destroy();
                }
            });
            this.charts = {};

            await this.loadStatsData();
        } finally {
            // 버튼 복원
            icon.className = 'fas fa-sync-alt me-1';
            btn.disabled = false;
        }
    }

    formatCurrency(amount) {
        if (!amount || amount === 0) return '0원';

        const numAmount = Math.abs(amount);
        let formatted;

        if (numAmount >= 100000000) {
            formatted = (numAmount / 100000000).toFixed(1) + '억';
        } else if (numAmount >= 10000) {
            formatted = (numAmount / 10000).toFixed(0) + '만';
        } else {
            formatted = numAmount.toLocaleString() + '원';
        }

        return amount < 0 ? '-' + formatted : formatted;
    }

    showLoading(show) {
        document.getElementById('loadingIndicator').style.display = show ? 'flex' : 'none';
    }

    showContent() {
        document.getElementById('statsContent').style.display = 'block';
    }
}

// 페이지 로드 시 초기화
window.statsManager = new StatsManager();
window.refreshStats = () => window.statsManager.refresh();

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        window.statsManager.init();
    });
} else {
    window.statsManager.init();
}
