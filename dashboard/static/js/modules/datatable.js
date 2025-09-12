// DataTable 관리 모듈
class DataTableManager {
    constructor() {
        this.dataTable = null;
        this.lockStatesRestored = false;
    }

    // DataTable 초기화
    initializeDataTable(data = []) {
        if (this.dataTable) {
            this.dataTable.destroy();
        }
        
        try {
            this.dataTable = $('#projectsTable').DataTable({
                data: data || [],
                columns: [
                    { data: '프로젝트 코드', width: '8%' },
                    { data: '담당자', width: '5%' },
                    { data: '거래처', width: '7%' },
                    { 
                        data: '현장 주소',
                        width: '26%',
                        render: function(data) {
                            if (data && data.trim() !== '') {
                                return data;
                            } else {
                                return '<i class="fas fa-exclamation-circle text-danger me-2" title="데이터가 비어 있습니다. 현장 주소를 입력해주세요." style="cursor: help;"></i><span class="text-muted">주소 정보 없음</span>';
                            }
                        }
                    },
                    { 
                        data: '공사 내용',
                        width: '19%',
                        render: function(data) {
                            if (data && data.trim() !== '') {
                                return data;
                            } else {
                                return '<i class="fas fa-exclamation-circle text-danger me-2" title="데이터가 비어 있습니다. 공사 내용을 입력해주세요." style="cursor: help;"></i><span class="text-muted">공사내용 없음</span>';
                            }
                        }
                    },
                    { 
                        data: '공사 시작',
                        width: '7%',
                        render: function(data) {
                            if (data) {
                                return formatDate(data);
                            } else {
                                return '<i class="fas fa-exclamation-triangle text-warning" title="데이터가 비어 있습니다. 공사 시작일을 입력해주세요." style="cursor: help; font-size: 1.2rem;"></i>';
                            }
                        }
                    },
                    { 
                        data: '공사 종료',
                        width: '7%',
                        render: function(data) {
                            if (data) {
                                return formatDate(data);
                            } else {
                                return '<i class="fas fa-exclamation-triangle text-warning" title="데이터가 비어 있습니다. 공사 종료일을 입력해주세요." style="cursor: help; font-size: 1.2rem;"></i>';
                            }
                        }
                    },
                    { 
                        data: null,
                        className: 'amount-cell',
                        width: '7%',
                        render: function(data, type, row) {
                            const totalAmount = row['총액 2'] || row['총액2'] || row['S'] || row['총액'] || 0;
                            const numValue = parseFloat(totalAmount) || 0;
                            
                            if (numValue > 0) {
                                return formatCurrency(numValue);
                            } else {
                                return '<i class="fas fa-exclamation-triangle text-warning" title="데이터가 비어 있습니다. 총액을 입력해주세요." style="cursor: help; font-size: 1.2rem;"></i>';
                            }
                        }
                    },
                    { 
                        data: null,
                        className: 'amount-cell',
                        width: '7%',
                        render: function(data, type, row) {
                            const outstandingData = row['미수금'] || row['미수금W'] || row['미수금 W'] || row['W'] || 0;
                            const outstandingAmount = parseFloat(outstandingData) || 0;
                            const totalAmount = parseFloat(row['총액 2'] || row['총액2'] || row['S'] || row['총액'] || 0);
                            
                            if (outstandingAmount >= 0) {
                                if (totalAmount === 0 && outstandingAmount === 0) {
                                    return '<i class="fas fa-exclamation-triangle text-warning" title="데이터가 비어 있습니다. 총액을 입력해주세요." style="cursor: help; font-size: 1.2rem;"></i>';
                                }
                                
                                if (outstandingAmount === 0) {
                                    return '<span class="text-muted">-</span>';
                                }
                                
                                return `<span class="amount-outstanding">${formatCurrency(outstandingAmount)}</span>`;
                            } else {
                                return '<span class="text-muted">-</span>';
                            }
                        }
                    },
                    { 
                        data: null,
                        width: '9%',
                        render: function(data, type, row) {
                            return getStatusBadge(row);
                        }
                    },
                    { 
                        data: null,
                        title: '데이터',
                        width: '4%',
                        className: 'text-center',
                        render: function(data, type, row) {
                            const completeness = checkDataCompleteness(row);
                            
                            if (completeness.missingCount === 0) {
                                return '<i class="fas fa-check-circle text-success" title="모든 필수 데이터가 입력되었습니다" style="font-size: 1.2rem;"></i>';
                            } else {
                                const missingList = completeness.missingFields.join(', ');
                                return `<i class="fas fa-exclamation-triangle text-warning" title="누락된 필드: ${missingList}" style="font-size: 1.2rem;"></i>`;
                            }
                        }
                    }
                ],
                pageLength: 25,
                processing: true,
                deferRender: true,
                scroller: false,
                stateSave: false,
                order: [[0, 'desc']],
                lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "전체"]],
                columnDefs: [
                    {
                        targets: 0,
                        type: 'num',
                        render: function(data, type, row) {
                            if (type === 'sort') {
                                return parseInt(data?.replace(/[^0-9]/g, '') || '0');
                            }
                            return `<span class="fw-bold" style="font-size: 1rem;">${createProjectBadge(data)}</span>`;
                        }
                    },
                    {
                        targets: 1,
                        render: function(data, type, row) {
                            if (type === 'sort') return data || '';
                            return createManagerBadge(data);
                        }
                    },
                    {
                        targets: 2,
                        render: function(data, type, row) {
                            if (type === 'sort') return data || '';
                            return createCompanyBadge(data);
                        }
                    }
                ],
                pageLength: 15,
                lengthMenu: [10, 15, 25, 50, 100],
                responsive: true,
                autoWidth: false,
                scrollX: true,
                scrollCollapse: true,
                fixedHeader: false,
                sScrollX: "100%",
                pagingType: "full_numbers",
                displayLength: 15,
                language: {
                    "emptyTable": "데이터가 없습니다.",
                    "info": "_START_~_END_ / 전체 _TOTAL_개 (페이지 _PAGE_ / _PAGES_)",
                    "infoEmpty": "0개 데이터",
                    "infoFiltered": "(전체 _MAX_개에서 필터링됨)",
                    "lengthMenu": "_MENU_개씩 보기",
                    "loadingRecords": "로딩 중...",
                    "processing": "처리 중...",
                    "search": "검색:",
                    "zeroRecords": "검색 결과가 없습니다.",
                    "paginate": {
                        "first": "처음",
                        "last": "마지막", 
                        "next": "다음",
                        "previous": "이전"
                    },
                    "aria": {
                        "sortAscending": ": 오름차순 정렬",
                        "sortDescending": ": 내림차순 정렬"
                    }
                },
                dom: '<"row"<"col-sm-12 col-md-6"l><"col-sm-12 col-md-6">>' +
                     '<"row"<"col-sm-12"tr>>' +
                     '<"row"<"col-sm-12 col-md-5"i><"col-sm-12 col-md-7"p>>',
                drawCallback: () => {
                    $('[data-bs-toggle="tooltip"]').tooltip();
                    
                    $('.dataTables_paginate .page-link').each(function() {
                        $(this).removeAttr('href');
                    });
                    
                    setTimeout(() => {
                        if (this.dataTable) {
                            this.dataTable.columns.adjust();
                            this.syncTableColumnWidths();
                        }
                    }, 5);
                    
                    setTimeout(() => {
                        this.customizePagination();
                    }, 10);
                    
                    this.setupRowClickEvents();
                    
                    if (!this.lockStatesRestored) {
                        this.lockStatesRestored = true;
                        setTimeout(() => {
                            restoreAllLockStates();
                        }, 100);
                    }
                }
            });
            
            setTimeout(() => {
                this.dataTable.columns.adjust();
                this.syncTableColumnWidths();
            }, 100);
            
            const rowCount = this.dataTable.data().length;
            const visibleRows = this.dataTable.rows().count();
            
            if (visibleRows === 0 && data && data.length > 0) {
                setTimeout(() => {
                    window.location.reload();
                }, 1000);
            }
            
        } catch (error) {
            // DataTable 생성 실패
            
            const tableBody = document.querySelector('#projectsTable tbody');
            if (tableBody && data && data.length > 0) {
                tableBody.innerHTML = '<tr><td colspan="10" class="text-center">테이블을 새로고침하고 있습니다...</td></tr>';
                
                setTimeout(() => {
                    window.location.reload();
                }, 1000);
            }
        }
    }

    // DataTable 헤더와 바디 컬럼 너비 동기화
    syncTableColumnWidths() {
        try {
            if (!this.dataTable) {
                return;
            }

            // 스크롤이 비활성화된 경우 기본 컬럼 조정만 수행
            this.dataTable.columns.adjust();
            
            // 스크롤 요소가 있는 경우에만 동기화 수행
            const headerTable = $('.dataTables_scrollHead table');
            const bodyTable = $('.dataTables_scrollBody table');
            
            if (headerTable.length && bodyTable.length) {
                const headerTableWidth = headerTable.outerWidth();
                const bodyTableWidth = bodyTable.outerWidth();
                const targetWidth = Math.min(headerTableWidth, bodyTableWidth);
                
                headerTable.css('width', targetWidth + 'px');
                bodyTable.css('width', targetWidth + 'px');
            }
        } catch (error) {
            // 컬럼 너비 동기화 실패 - 무시
        }
    }

    // 페이지네이션 커스터마이즈
    customizePagination() {
        const paginateContainer = $('.dataTables_paginate');
        const pageButtons = paginateContainer.find('.paginate_button');
        
        if (pageButtons.length > 9) {
            const currentPageButton = paginateContainer.find('.current');
            const currentPageIndex = pageButtons.index(currentPageButton);
            
            pageButtons.hide();
            
            paginateContainer.find('.first').show();
            paginateContainer.find('.previous').show();
            
            let startIndex = Math.max(1, currentPageIndex - 2);
            let endIndex = Math.min(pageButtons.length - 2, currentPageIndex + 2);
            
            if (endIndex - startIndex < 4) {
                if (startIndex === 1) {
                    endIndex = Math.min(pageButtons.length - 2, startIndex + 4);
                } else if (endIndex === pageButtons.length - 2) {
                    startIndex = Math.max(1, endIndex - 4);
                }
            }
            
            for (let i = startIndex; i <= endIndex; i++) {
                $(pageButtons[i]).show();
            }
            
            paginateContainer.find('.next').show();
            paginateContainer.find('.last').show();
        }
    }

    // 행 클릭 이벤트 설정
    setupRowClickEvents() {
        $('#projectsTable tbody tr').off('click').on('click', function(e) {
            if ($(e.target).closest('button, a, .btn').length > 0) {
                return;
            }
            
            const rowData = window.dataTableManager.dataTable.row(this).data();
            if (rowData && rowData['프로젝트 코드']) {
                viewProject(rowData['프로젝트 코드']);
            }
        });
    }

    // DataTable 인스턴스 반환
    getDataTable() {
        return this.dataTable;
    }
}

// 전역 DataTableManager 인스턴스 생성
window.dataTableManager = new DataTableManager();

// 전역 함수들을 위한 래퍼
function initializeDataTable(data) {
    return window.dataTableManager.initializeDataTable(data);
}

function syncTableColumnWidths() {
    return window.dataTableManager.syncTableColumnWidths();
}

function customizePagination() {
    return window.dataTableManager.customizePagination();
}

function setupRowClickEvents() {
    return window.dataTableManager.setupRowClickEvents();
}