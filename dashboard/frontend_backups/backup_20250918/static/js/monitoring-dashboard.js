/**
 * 성능 모니터링 대시보드 JavaScript
 * 실시간 데이터 업데이트, 차트 관리, 알림 처리
 */

class MonitoringDashboard {
    constructor() {
        this.charts = {};
        this.refreshInterval = null;
        this.wsConnection = null;
        this.dataHistory = {
            cpu: [],
            memory: [],
            disk: [],
            responseTime: [],
            errorCount: []
        };
        this.maxHistoryPoints = 50;
        this.refreshRate = 30000; // 30초

        this.init();
    }

    /**
     * 대시보드 초기화
     */
    init() {
        this.initializeCharts();
        this.setupEventListeners();
        this.startDataCollection();

        // WebSocket 연결 시도 (실시간 업데이트용)
        this.initWebSocket();

        console.log('📊 모니터링 대시보드가 초기화되었습니다.');
    }

    /**
     * 차트 초기화
     */
    initializeCharts() {
        // HTTP 성능 트렌드 차트
        this.initHttpPerformanceChart();

        // 시스템 리소스 차트
        this.initSystemResourceChart();

        // 응답시간 히스토리 차트
        this.initResponseTimeChart();

        // 에러 트렌드 차트
        this.initErrorTrendChart();
    }

    /**
     * HTTP 성능 트렌드 차트 초기화
     */
    initHttpPerformanceChart() {
        const ctx = document.getElementById('httpPerformanceChart');
        if (!ctx) return;

        this.charts.httpPerformance = new Chart(ctx, {
            type: 'line',
            data: {
                labels: [],
                datasets: [{
                    label: '평균 응답시간 (ms)',
                    data: [],
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }, {
                    label: '초당 요청수 (RPS)',
                    data: [],
                    borderColor: '#f093fb',
                    backgroundColor: 'rgba(240, 147, 251, 0.1)',
                    tension: 0.4,
                    yAxisID: 'y1',
                    pointRadius: 4,
                    pointHoverRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: {
                    intersect: false,
                    mode: 'index'
                },
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            padding: 20
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        titleColor: 'white',
                        bodyColor: 'white',
                        borderColor: '#667eea',
                        borderWidth: 1
                    }
                },
                scales: {
                    x: {
                        grid: {
                            color: 'rgba(0,0,0,0.1)'
                        },
                        title: {
                            display: true,
                            text: '시간'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0,0,0,0.1)'
                        },
                        title: {
                            display: true,
                            text: '응답시간 (ms)'
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'RPS'
                        },
                        grid: {
                            drawOnChartArea: false
                        }
                    }
                },
                animation: {
                    duration: 750,
                    easing: 'easeInOutQuart'
                }
            }
        });
    }

    /**
     * 시스템 리소스 차트 초기화
     */
    initSystemResourceChart() {
        const ctx = document.getElementById('systemResourceChart');
        if (!ctx) return;

        this.charts.systemResource = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['CPU 사용률', '메모리 사용률', '디스크 사용률'],
                datasets: [{
                    data: [0, 0, 0],
                    backgroundColor: [
                        'rgba(255, 99, 132, 0.8)',
                        'rgba(54, 162, 235, 0.8)',
                        'rgba(255, 206, 86, 0.8)'
                    ],
                    borderColor: [
                        'rgba(255, 99, 132, 1)',
                        'rgba(54, 162, 235, 1)',
                        'rgba(255, 206, 86, 1)'
                    ],
                    borderWidth: 2,
                    hoverOffset: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '60%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 20,
                            usePointStyle: true
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return `${context.label}: ${context.parsed}%`;
                            }
                        }
                    }
                },
                animation: {
                    animateScale: true,
                    animateRotate: true
                }
            }
        });
    }

    /**
     * 응답시간 히스토리 차트 초기화
     */
    initResponseTimeChart() {
        const ctx = document.getElementById('responseTimeChart');
        if (!ctx) return;

        this.charts.responseTime = new Chart(ctx, {
            type: 'line',
            data: {
                labels: [],
                datasets: [{
                    label: '응답시간 (ms)',
                    data: [],
                    borderColor: '#36a2eb',
                    backgroundColor: 'rgba(54, 162, 235, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 2,
                    pointHoverRadius: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    x: { display: false },
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(0,0,0,0.1)' }
                    }
                }
            }
        });
    }

    /**
     * 에러 트렌드 차트 초기화
     */
    initErrorTrendChart() {
        const ctx = document.getElementById('errorTrendChart');
        if (!ctx) return;

        this.charts.errorTrend = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: [],
                datasets: [{
                    label: '에러 수',
                    data: [],
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false }
                },
                scales: {
                    x: { display: false },
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(0,0,0,0.1)' }
                    }
                }
            }
        });
    }

    /**
     * 이벤트 리스너 설정
     */
    setupEventListeners() {
        // 새로고침 버튼
        const refreshBtn = document.querySelector('.refresh-btn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => this.refreshAllData());
        }

        // 자동 새로고침 토글
        const autoRefreshToggle = document.getElementById('autoRefreshToggle');
        if (autoRefreshToggle) {
            autoRefreshToggle.addEventListener('change', (e) => {
                if (e.target.checked) {
                    this.startAutoRefresh();
                } else {
                    this.stopAutoRefresh();
                }
            });
        }

        // 새로고침 간격 설정
        const refreshIntervalSelect = document.getElementById('refreshInterval');
        if (refreshIntervalSelect) {
            refreshIntervalSelect.addEventListener('change', (e) => {
                this.refreshRate = parseInt(e.target.value) * 1000;
                this.restartAutoRefresh();
            });
        }

        // 페이지 가시성 변경 시 자동 새로고침 제어
        document.addEventListener('visibilitychange', () => {
            if (document.visibilityState === 'visible') {
                this.startAutoRefresh();
            } else {
                this.stopAutoRefresh();
            }
        });

        // 윈도우 리사이즈 시 차트 리사이즈
        window.addEventListener('resize', this.debounce(() => {
            Object.values(this.charts).forEach(chart => {
                if (chart && typeof chart.resize === 'function') {
                    chart.resize();
                }
            });
        }, 300));
    }

    /**
     * WebSocket 연결 초기화
     */
    initWebSocket() {
        if (typeof io !== 'undefined') {
            try {
                this.wsConnection = io();

                this.wsConnection.on('connect', () => {
                    console.log('🔗 WebSocket 연결됨');
                    this.showNotification('실시간 모니터링이 활성화되었습니다.', 'success');
                });

                this.wsConnection.on('disconnect', () => {
                    console.log('🔌 WebSocket 연결 해제됨');
                    this.showNotification('실시간 연결이 해제되었습니다.', 'warning');
                });

                this.wsConnection.on('metric_update', (data) => {
                    this.handleRealtimeMetricUpdate(data);
                });

                this.wsConnection.on('error_alert', (data) => {
                    this.handleErrorAlert(data);
                });

            } catch (error) {
                console.warn('WebSocket 연결 실패:', error);
            }
        }
    }

    /**
     * 실시간 메트릭 업데이트 처리
     */
    handleRealtimeMetricUpdate(data) {
        if (data.type === 'system') {
            this.updateSystemMetrics(data.metrics);
        } else if (data.type === 'http') {
            this.updateHttpMetrics(data.metrics);
        }
    }

    /**
     * 에러 알림 처리
     */
    handleErrorAlert(alert) {
        this.showNotification(
            `⚠️ ${alert.title}`,
            'danger',
            5000
        );

        // 에러 카운터 업데이트
        this.incrementErrorCounter();
    }

    /**
     * 데이터 수집 시작
     */
    startDataCollection() {
        this.refreshAllData();
        this.startAutoRefresh();
    }

    /**
     * 자동 새로고침 시작
     */
    startAutoRefresh() {
        this.stopAutoRefresh(); // 기존 인터벌 정리

        this.refreshInterval = setInterval(() => {
            this.refreshAllData();
        }, this.refreshRate);

        console.log(`🔄 자동 새로고침 시작 (${this.refreshRate/1000}초 간격)`);
    }

    /**
     * 자동 새로고침 중지
     */
    stopAutoRefresh() {
        if (this.refreshInterval) {
            clearInterval(this.refreshInterval);
            this.refreshInterval = null;
        }
    }

    /**
     * 자동 새로고침 재시작
     */
    restartAutoRefresh() {
        this.stopAutoRefresh();
        this.startAutoRefresh();
    }

    /**
     * 전체 데이터 새로고침
     */
    async refreshAllData() {
        this.updateLastUpdatedTime();
        this.showLoadingState(true);

        try {
            const loadPromises = [
                this.loadSystemMetrics(),
                this.loadHttpMetrics(),
                this.loadErrorMetrics(),
                this.loadDetailedMetrics()
            ];

            await Promise.allSettled(loadPromises);

            this.showNotification('데이터가 업데이트되었습니다.', 'success', 2000);

        } catch (error) {
            console.error('데이터 새로고침 실패:', error);
            this.showNotification('데이터 새로고침 중 오류가 발생했습니다.', 'danger');
        } finally {
            this.showLoadingState(false);
        }
    }

    /**
     * 시스템 메트릭 로드
     */
    async loadSystemMetrics() {
        try {
            const response = await fetch('/api/system/stats');
            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const data = await response.json();
            this.updateSystemMetrics(data);

        } catch (error) {
            console.error('시스템 메트릭 로딩 실패:', error);
            this.showMetricError('시스템 메트릭');
        }
    }

    /**
     * HTTP 메트릭 로드
     */
    async loadHttpMetrics() {
        try {
            const response = await fetch('/api/metrics/query?name=http.request.duration&limit=30');
            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const data = await response.json();
            this.updateHttpMetrics(data);

        } catch (error) {
            console.error('HTTP 메트릭 로딩 실패:', error);
            this.showMetricError('HTTP 성능');
        }
    }

    /**
     * 에러 메트릭 로드
     */
    async loadErrorMetrics() {
        try {
            const response = await fetch('/api/errors/dashboard');
            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const data = await response.json();
            this.updateErrorMetrics(data);

        } catch (error) {
            console.error('에러 메트릭 로딩 실패:', error);
            this.showMetricError('에러 현황');
        }
    }

    /**
     * 상세 메트릭 로드
     */
    async loadDetailedMetrics() {
        try {
            const response = await fetch('/api/metrics/summary');
            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const data = await response.json();
            this.updateDetailedMetrics(data);

        } catch (error) {
            console.error('상세 메트릭 로딩 실패:', error);
            this.showMetricError('상세 메트릭');
        }
    }

    /**
     * 시스템 메트릭 업데이트
     */
    updateSystemMetrics(data) {
        // CPU 사용률
        this.updateMetricDisplay('cpuUsage', data.cpu.percent, '%');
        this.addToHistory('cpu', data.cpu.percent);

        // 메모리 사용률
        this.updateMetricDisplay('memoryUsage', data.memory.percent, '%');
        this.addToHistory('memory', data.memory.percent);

        // 디스크 사용률
        this.updateMetricDisplay('diskUsage', data.disk.percent, '%');
        this.addToHistory('disk', data.disk.percent);

        // 시스템 리소스 차트 업데이트
        if (this.charts.systemResource) {
            this.charts.systemResource.data.datasets[0].data = [
                data.cpu.percent,
                data.memory.percent,
                data.disk.percent
            ];
            this.charts.systemResource.update('none');
        }
    }

    /**
     * HTTP 메트릭 업데이트
     */
    updateHttpMetrics(data) {
        if (!data.metrics || data.metrics.length === 0) return;

        // 평균 응답시간 계산
        const avgResponseTime = data.metrics.reduce((sum, m) => sum + m.value, 0) / data.metrics.length;
        this.updateMetricDisplay('responseTime', avgResponseTime, 'ms');
        this.addToHistory('responseTime', avgResponseTime);

        // HTTP 성능 차트 업데이트
        if (this.charts.httpPerformance) {
            const labels = data.metrics.map(m =>
                new Date(m.timestamp).toLocaleTimeString('ko-KR', {
                    hour: '2-digit',
                    minute: '2-digit'
                })
            );
            const values = data.metrics.map(m => m.value);

            this.charts.httpPerformance.data.labels = labels.slice(-20);
            this.charts.httpPerformance.data.datasets[0].data = values.slice(-20);
            this.charts.httpPerformance.update('none');
        }
    }

    /**
     * 에러 메트릭 업데이트
     */
    updateErrorMetrics(data) {
        const errorStatusDiv = document.getElementById('errorStatus');
        if (!errorStatusDiv) return;

        const totalErrors = data.summary.errors_last_hour || 0;

        if (totalErrors === 0) {
            errorStatusDiv.innerHTML = `
                <div class="text-success text-center">
                    <i class="fas fa-check-circle fa-3x mb-3"></i>
                    <h6>시스템 정상</h6>
                    <p class="small text-muted">최근 1시간 동안 에러 없음</p>
                </div>
            `;
        } else {
            const severity = totalErrors > 10 ? 'danger' : 'warning';
            errorStatusDiv.innerHTML = `
                <div class="text-${severity} text-center">
                    <i class="fas fa-exclamation-triangle fa-3x mb-3"></i>
                    <h6>에러 감지</h6>
                    <p class="small">최근 1시간: <strong>${totalErrors}건</strong></p>
                    <button class="btn btn-sm btn-outline-${severity}" onclick="window.open('/monitoring/errors', '_blank')">
                        상세 보기
                    </button>
                </div>
            `;
        }

        // 에러 히스토리 업데이트
        this.addToHistory('errorCount', totalErrors);
        this.updateErrorTrendChart();
    }

    /**
     * 상세 메트릭 업데이트
     */
    updateDetailedMetrics(data) {
        const tbody = document.querySelector('#metricsTable tbody');
        if (!tbody || !data.top_metrics) return;

        tbody.innerHTML = data.top_metrics.slice(0, 10).map(metric => `
            <tr>
                <td>
                    <code class="text-primary">${metric.name}</code>
                </td>
                <td>
                    <span class="badge bg-info">${metric.count.toLocaleString()}</span>
                </td>
                <td>
                    <span class="badge bg-secondary">counter</span>
                </td>
                <td>개</td>
                <td>
                    <small class="text-muted">${new Date().toLocaleTimeString()}</small>
                </td>
                <td>
                    <i class="fas fa-check-circle text-success" title="정상"></i>
                </td>
            </tr>
        `).join('');
    }

    /**
     * 메트릭 표시 업데이트
     */
    updateMetricDisplay(elementId, value, unit) {
        const element = document.getElementById(elementId);
        if (!element) return;

        const formattedValue = typeof value === 'number' ?
            (value < 10 ? value.toFixed(2) : value.toFixed(1)) : value;

        element.textContent = `${formattedValue}${unit}`;

        // 상태에 따른 색상 변경
        const status = this.getMetricStatus(value, elementId);
        element.className = `metric-value ${status}`;

        // 트렌드 표시 업데이트
        this.updateTrendDisplay(elementId, value);
    }

    /**
     * 히스토리에 데이터 추가
     */
    addToHistory(type, value) {
        if (!this.dataHistory[type]) {
            this.dataHistory[type] = [];
        }

        this.dataHistory[type].push({
            value: value,
            timestamp: new Date()
        });

        // 최대 점수 제한
        if (this.dataHistory[type].length > this.maxHistoryPoints) {
            this.dataHistory[type].shift();
        }
    }

    /**
     * 트렌드 표시 업데이트
     */
    updateTrendDisplay(elementId, currentValue) {
        const trendElementId = elementId.replace('Usage', 'Trend').replace('Time', 'Trend');
        const trendElement = document.getElementById(trendElementId);
        if (!trendElement) return;

        const type = elementId.replace('Usage', '').replace('responseTime', 'responseTime').toLowerCase();
        const history = this.dataHistory[type];

        if (!history || history.length < 2) {
            trendElement.innerHTML = '<i class="fas fa-minus me-1"></i>데이터 수집 중...';
            return;
        }

        const previousValue = history[history.length - 2].value;
        const change = currentValue - previousValue;
        const changePercent = previousValue !== 0 ? (change / previousValue) * 100 : 0;

        let trendClass, trendIcon, trendText;

        if (Math.abs(changePercent) < 1) {
            trendClass = 'trend-stable';
            trendIcon = 'fas fa-minus';
            trendText = '안정';
        } else if (change > 0) {
            trendClass = type === 'responseTime' ? 'trend-down' : 'trend-up';
            trendIcon = 'fas fa-arrow-up';
            trendText = `+${changePercent.toFixed(1)}%`;
        } else {
            trendClass = type === 'responseTime' ? 'trend-up' : 'trend-down';
            trendIcon = 'fas fa-arrow-down';
            trendText = `${changePercent.toFixed(1)}%`;
        }

        trendElement.className = `metric-trend ${trendClass}`;
        trendElement.innerHTML = `<i class="${trendIcon} me-1"></i>${trendText}`;
    }

    /**
     * 에러 트렌드 차트 업데이트
     */
    updateErrorTrendChart() {
        if (!this.charts.errorTrend || !this.dataHistory.errorCount) return;

        const labels = this.dataHistory.errorCount.map((_, index) => `${index + 1}`);
        const values = this.dataHistory.errorCount.map(item => item.value);

        this.charts.errorTrend.data.labels = labels.slice(-10);
        this.charts.errorTrend.data.datasets[0].data = values.slice(-10);
        this.charts.errorTrend.update('none');
    }

    /**
     * 메트릭 상태 결정
     */
    getMetricStatus(value, type) {
        if (type.includes('cpu') || type.includes('memory') || type.includes('disk')) {
            if (value >= 90) return 'status-critical';
            if (value >= 70) return 'status-warning';
            return 'status-healthy';
        }

        if (type.includes('responseTime')) {
            if (value >= 1000) return 'status-critical';
            if (value >= 500) return 'status-warning';
            return 'status-healthy';
        }

        return 'status-healthy';
    }

    /**
     * 마지막 업데이트 시간 갱신
     */
    updateLastUpdatedTime() {
        const element = document.getElementById('lastUpdated');
        if (element) {
            element.textContent = `마지막 업데이트: ${new Date().toLocaleString('ko-KR')}`;
        }
    }

    /**
     * 로딩 상태 표시
     */
    showLoadingState(isLoading) {
        const refreshBtn = document.querySelector('.refresh-btn');
        if (!refreshBtn) return;

        if (isLoading) {
            refreshBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>업데이트 중...';
            refreshBtn.disabled = true;
        } else {
            refreshBtn.innerHTML = '<i class="fas fa-sync-alt me-2"></i>새로고침';
            refreshBtn.disabled = false;
        }
    }

    /**
     * 메트릭 에러 표시
     */
    showMetricError(metricType) {
        console.warn(`${metricType} 로딩 실패 - 인증이 필요할 수 있습니다.`);
    }

    /**
     * 알림 표시
     */
    showNotification(message, type = 'info', duration = 3000) {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
        alertDiv.style.cssText = `
            top: 20px;
            right: 20px;
            z-index: 9999;
            min-width: 300px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        `;

        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;

        document.body.appendChild(alertDiv);

        // 자동 제거
        setTimeout(() => {
            if (alertDiv.parentNode) {
                alertDiv.remove();
            }
        }, duration);
    }

    /**
     * 에러 카운터 증가
     */
    incrementErrorCounter() {
        const errorCount = this.dataHistory.errorCount;
        const currentCount = errorCount.length > 0 ? errorCount[errorCount.length - 1].value : 0;
        this.addToHistory('errorCount', currentCount + 1);
        this.updateErrorTrendChart();
    }

    /**
     * 디바운스 유틸리티
     */
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    /**
     * 대시보드 정리
     */
    destroy() {
        this.stopAutoRefresh();

        if (this.wsConnection) {
            this.wsConnection.disconnect();
        }

        Object.values(this.charts).forEach(chart => {
            if (chart && typeof chart.destroy === 'function') {
                chart.destroy();
            }
        });

        console.log('📊 모니터링 대시보드가 정리되었습니다.');
    }
}

// 전역 대시보드 인스턴스
let dashboard;

// DOM 로드 완료 시 대시보드 초기화
document.addEventListener('DOMContentLoaded', function() {
    dashboard = new MonitoringDashboard();
});

// 페이지 언로드 시 정리
window.addEventListener('beforeunload', function() {
    if (dashboard) {
        dashboard.destroy();
    }
});

// 전역 함수 (HTML에서 호출용)
function refreshAllData() {
    if (dashboard) {
        dashboard.refreshAllData();
    }
}