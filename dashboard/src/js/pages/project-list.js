/**
 * Project List Page - 메인 진입점 (Orchestration Only)
 * 전문가 리뷰: "app.init()이 orchestration만 담당하게 리팩터링"
 */

// 서비스 임포트 (데이터, 상태, 캐시 관리 분리)
import DataManager from '../services/DataManager.js';
import StateManager from '../services/StateManager.js';
import CacheStatusManager from '../services/CacheStatusManager.js';
import MetadataManager from '../services/MetadataManager.js';

// 유틸리티 임포트
import { buttonStateManager } from '../utils/ButtonStateManager.js';
import DataValidator from '../utils/DataValidator.js';
import { exposeGlobalAPI } from '../api/ProjectListAPI.js';
import { addSkipLink, enhanceTableAccessibility, announceToScreenReader } from '../utils/accessibility.js';

import logger from '../utils/logger.js';
// 컴포넌트 임포트 (동적 로딩으로 성능 최적화)
const ProjectTable = () => import('../components/ProjectTable.js');
const ModernProjectFilters = () => import('../components/ModernProjectFilters.js');
const ModernProjectModal = () => import('../components/ModernProjectModal.js');
const UserModal = () => import('../components/UserModal.js');
const AuditLogModal = () => import('../components/AuditLogModal.js');
const MobileCardView = () => import('../components/MobileCardView.js');
const ProjectRowAccordion = () => import('../components/ProjectRowAccordion.js');
const AsView = () => import('../components/AsView.js');

/**
 * 프로젝트 리스트 페이지 메인 클래스 (Orchestration Only)
 * 전문가 리뷰: "데이터 로딩, UI 컴포넌트, 상태 저장을 각각 별도 클래스로 나눠"
 */
class ProjectListApp {
  constructor() {
    // 서비스 인스턴스 (분리된 책임)
    this.dataManager = new DataManager();
    this.stateManager = new StateManager();
    this.metadataManager = new MetadataManager();
    this.cacheStatusManager = null; // 소켓 연결 후 초기화

    // 컴포넌트 컨테이너
    this.components = {};
    this.isInitialized = false;
    this.socket = null;
    this.isRefreshing = false; // 중복 새로고침 방지 플래그
    this.eventListenersAttached = false; // 이벤트 리스너 중복 등록 방지 플래그
  }

  /**
   * 애플리케이션 초기화 (Orchestration Only)
   * 전문가 리뷰: "app.init()이 orchestration만 담당하게"
   */
  async init() {
    logger.debug('[ProjectListApp] init() 호출됨, isInitialized:', this.isInitialized);
    console.trace('[ProjectListApp] init() 호출 스택');

    if (this.isInitialized) {
      logger.warn('[ProjectListApp] 이미 초기화되었습니다. 중복 init() 호출 방지');
      return;
    }

    // 즉시 플래그 설정하여 중복 호출 방지
    this.isInitialized = true;
    logger.debug('[ProjectListApp] isInitialized 플래그 설정됨');

    try {
      logger.debug('[ProjectListApp] Orchestration 시작');

      // 0. 마지막 visibility 변경 시간 초기화
      this.lastVisibilityChange = Date.now();

      // 1. 메타데이터 로드 (사용자 권한, 기본 설정)
      const metadata = this.metadataManager.loadPageMetadata();
      logger.debug('[ProjectListApp] 페이지 메타데이터 로드:', metadata);

      // 2. 상태 관리 초기화 (메타데이터 기반)
      this.stateManager.initializeState();
      this.stateManager.setUserPermissions(this.metadataManager.getUserPermissions());

      // 3. 웹소켓 연결 설정
      this.setupWebSocketConnection();

      // 4. 캐시 상태 관리 초기화 (서버 푸시 기반)
      this.cacheStatusManager = new CacheStatusManager(this.socket);
      this.cacheStatusManager.exposeGlobalFunction();
      this.cacheStatusManager.startMonitoring();

      // 5. 컴포넌트 로딩 및 등록
      await this.loadComponents();
      this.stateManager.registerComponents(this.components);

      // 6. 이벤트 리스너 설정 (단순화) - 중복 등록 방지 포함
      this.setupEventListeners();

      // 7. 초기 데이터 로드 (DataManager 사용)
      await this.loadInitialData();

      // 8. 통계 데이터 별도 로드 (API 기반)
      await this.loadStatisticsData();

      // 9. 권한 기반 UI 적용
      this.stateManager.applyPermissions();

      // 10. 접근성(A11y) 개선 적용
      this.enhanceAccessibility();

      logger.debug('[ProjectListApp] Orchestration 완료');

    } catch (error) {
      logger.error('[ProjectListApp] 초기화 실패:', error);
      this.showErrorMessage('애플리케이션 초기화에 실패했습니다.');
      this.isInitialized = false; // 실패 시 플래그 재설정
    }
  }

  /**
   * 컴포넌트 동적 로딩
   */
  async loadComponents() {

    const loadPromises = [
      ProjectTable().then(module => {
        this.components.table = new module.default();
      }),
      ModernProjectFilters().then(module => {
        this.components.filters = new module.default();
      }),
      ModernProjectModal().then(module => {
        this.components.modernModal = new module.default();
        // 옵션 데이터 프리페치 (백그라운드)
        this.components.modernModal.prefetchOptions();
      }).catch(err => {
        logger.error('[ERROR] ModernProjectModal 로드 실패:', err);
      }),
      UserModal().then(module => {
        this.components.userModal = new module.default();
      }),
      AuditLogModal().then(module => {
        this.components.auditLogModal = new module.default();
      }).catch(err => {
        logger.error('[ERROR] AuditLogModal 로드 실패:', err);
      }),
      MobileCardView().then(module => {
        this.components.mobileView = new module.default();
      }),
      // [ROCKET] 핵심 컴포넌트들 초기화 추가
      ProjectRowAccordion().then(module => {
        this.components.accordion = new module.default();
      }),
      AsView().then(module => {
        this.components.asView = new module.default();
        window.asView = this.components.asView;  // accordion 'A/S 요청' 버튼에서 접근
        // byCode 는 A/S 관리 모드 진입 시에만 로드 (일반 모드는 'A/S 요청'만 노출).
      }).catch(err => {
        logger.error('[ERROR] AsView 로드 실패:', err);
      })
    ];

    await Promise.all(loadPromises);

    // 현대화된 필터 컴포넌트 초기화 및 콜백 등록
    if (this.components.filters) {
      await this.components.filters.init();
      this.components.filters.onFilterChange((filteredData) => {
        this.filteredData = filteredData;
        this.components.table?.updateData(filteredData);
        this.components.mobileView?.updateCards(filteredData);
      });

      // HTML에서 접근할 수 있도록 글로벌 변수 설정
      window.modernFilters = this.components.filters;
    }

    // [ROCKET] 테이블은 updateUI()에서 데이터와 함께 초기화됨
    logger.debug('[SUCCESS] 테이블 컴포넌트 로드 완료 (초기화는 updateUI에서 수행)');

    // [ROCKET] 액션 버튼 초기화
    if (this.components.actionButtons) {
      await this.components.actionButtons.init();
    }

  }

  /**
   * 이벤트 리스너 설정
   */
  setupEventListeners() {
    // 새로고침 버튼
    const refreshBtn = document.getElementById('refreshProjectBtn');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => this.refreshData(true, true, true)); // forceRefresh=true, showOverlay=true, showMessage=true
    }

    // 글로벌 키보드 단축키
    document.addEventListener('keydown', (e) => {
      if (e.ctrlKey || e.metaKey) {
        switch (e.key) {
          case 'r':
            e.preventDefault();
            this.refreshData(true, true, true); // forceRefresh=true, showOverlay=true, showMessage=true
            break;
          case 'f':
            e.preventDefault();
            this.components.filters?.focusSearch();
            break;
        }
      }
    });

    // 웹소켓 연결은 init()에서 이미 수행되므로 여기서는 제거 (중복 방지)

    // 프로젝트 모달에서 발생하는 projectUpdated 이벤트 처리 (중복 등록 방지)
    if (!this.eventListenersAttached) {
      window.addEventListener('projectUpdated', (event) => {
        const { partialUpdate, project, projectCode, action } = event.detail || {};

        logger.debug(`🔔 [이벤트] projectUpdated 수신:`, {
          partialUpdate,
          hasProject: !!project,
          projectCode,
          action
        });

        // 부분 업데이트 플래그가 있으면 UI 업데이트 건너뛰기
        if (partialUpdate) {
          logger.debug(`🔄 [부분 업데이트] 프로젝트 ${projectCode} - UI 업데이트 건너뜀 (백엔드 캐시만 무효화)`);

          // 캐시만 무효화 (다음 로드 시 최신 데이터 가져오기 위해)
          this.dataManager.clearCache();

          // UI는 전혀 건드리지 않음 - 아코디언이 열려있으면 그대로 유지
          logger.debug(`✅ [캐시 무효화] 다음 데이터 로드 시 최신 데이터 사용`);

          // 필터 결과에서 사라진 프로젝트의 아코디언 자동 정리 (수금 관리 모드에서 자주 발생)
          // 예:
          //   · 미수금 필터 ON + 미수금=0 저장 → 필터 outstanding 에서 제외
          //   · 미수금 필터 ON + 공사 취소 → '수금 관련 특이사항' 에 '공사 취소' 추가되어 자동 제외
          //   · 미수금 필터 ON + '수금 확인' 체크 → 미수금 처리 완료로 제외
          try {
            const accordion = window.projectListApp?.components?.accordion;
            const filtered = this.stateManager?.filteredData || [];
            const stillVisible = filtered.some(p => p?.['프로젝트 코드'] === projectCode);
            if (!stillVisible && accordion?.isOpen &&
                accordion?.currentProject?.['프로젝트 코드'] === projectCode) {
              logger.info(`[필터-아코디언] ${projectCode} ${action || 'edit'} 후 필터 결과에서 제외됨 → 아코디언 자동 닫기`);
              // 편집 모드였다면 정리 (편집 상태에서 사라지는 건 이미 저장 완료 후라 안전)
              if (accordion.modeManager?.isEditMode?.()) {
                accordion.disableUnifiedEditMode(projectCode);
              }
              accordion.closeAccordion?.();
              // action 별 안내 메시지 세분화
              let msg;
              if (action === 'cancel_construction') {
                msg = `${projectCode} 공사 취소 완료. 취소된 공사는 수금 관리 모드에서 자동으로 숨겨집니다. 다시 확인하려면 상단 '수금 관리 모드'를 끄세요.`;
              } else if (action === 'resume_construction') {
                msg = `${projectCode} 공사 재개 완료.`;
              } else {
                msg = `${projectCode} 저장 완료. 현재 필터 조건에 맞지 않아 리스트에서 제외되었습니다.`;
              }
              if (window.showSystemAlert) {
                window.showSystemAlert(msg, 'info');
              }
            }
          } catch (err) {
            logger.warn('[필터-아코디언] 정리 로직 오류 (무시):', err);
          }
        } else {
          // 전체 새로고침 (레거시 방식 또는 partialUpdate 플래그 없을 때)
          logger.warn(`⚠️ [전체 새로고침] partialUpdate 플래그 없음 - 호출 위치 확인 필요`);
          console.trace('[전체 새로고침 스택]');
          this.dataManager.clearCache();
          this.refreshData(true, false); // 오버레이 없이 새로고침
        }
      });

      this.eventListenersAttached = true;
      logger.debug('✅ [이벤트] projectUpdated 이벤트 리스너 등록 완료');
    }

    // [ROCKET] LocalStorage 브로드캐스트 이벤트 (탭 간 동기화)
    window.addEventListener('storage', (event) => {
      // 새 프로젝트 생성 브로드캐스트
      if (event.key === 'newProjectCreated' && event.newValue) {
        try {
          const newProjectData = JSON.parse(event.newValue);
          logger.debug('다른 탭에서 새 프로젝트 생성됨:', newProjectData);
          // DataManager를 통한 캐시 클리어 및 데이터 새로고침
          this.dataManager.clearCache();
          this.refreshData(true, false); // 오버레이 없이 새로고침
          // 처리 완료 후 LocalStorage 항목 제거
          localStorage.removeItem('newProjectCreated');
        } catch (error) {
          logger.error('newProjectCreated 이벤트 처리 오류:', error);
        }
      }

    });
  }

  /**
   * 로딩 오버레이 표시
   */
  showLoadingOverlay() {
    const overlay = document.getElementById('tableLoadingOverlay');
    const tableWrapper = document.querySelector('.table-wrapper');
    const mobileContainer = document.getElementById('mobileCardContainer');

    // 오버레이 표시
    if (overlay) {
      overlay.style.display = 'flex';
    }

    // 테이블 래퍼 및 모바일 컨테이너 숨김 (visibility 사용)
    if (tableWrapper) {
      tableWrapper.style.visibility = 'hidden';
      tableWrapper.style.opacity = '0';
    }
    if (mobileContainer) {
      mobileContainer.style.visibility = 'hidden';
      mobileContainer.style.opacity = '0';
    }
  }

  /**
   * 로딩 오버레이 숨김 및 테이블 표시
   */
  hideLoadingOverlay() {
    const overlay = document.getElementById('tableLoadingOverlay');
    const tableWrapper = document.querySelector('.table-wrapper');
    const mobileContainer = document.getElementById('mobileCardContainer');

    // 오버레이 숨김
    if (overlay) {
      overlay.style.display = 'none';
    }

    // 테이블 래퍼 및 모바일 컨테이너 표시 (헤더와 데이터 동시에 나타남)
    if (tableWrapper) {
      tableWrapper.style.visibility = 'visible';
      tableWrapper.style.opacity = '1';
    }
    if (mobileContainer) {
      mobileContainer.style.visibility = 'visible';
      mobileContainer.style.opacity = '1';
    }
  }

  /**
   * 초기 데이터 로드 (DataManager 사용)
   */
  async loadInitialData() {
    try {
      logger.debug('[ProjectListApp] 초기 데이터 로드 시작');

      // 로딩 오버레이 표시 (HTML에서는 display: none으로 시작)
      this.showLoadingOverlay();

      // DataManager를 통한 데이터 로드
      const result = await this.dataManager.loadProjectData();

      if (result.data && result.data.length > 0) {
        // StateManager를 통한 상태 업데이트 (필터 재적용 건너뜀 - updateUI에서 처리)
        this.stateManager.setCurrentData(result.data, true);

        // UI 업데이트 (StateManager의 리스너 통해 자동 처리)
        await this.updateUI(result.data);

        logger.debug(`[ProjectListApp] 데이터 로드 완료: ${result.data.length}개 항목 (${result.fromCache ? '캐시' : '서버'})`);

        // DataTable draw 이벤트 한 번만 리스닝하여 렌더링 완료 감지
        const table = this.components.table?.table;
        if (table) {
          table.one('draw.dt', () => {
            logger.debug('[ProjectListApp] DataTable 렌더링 완료');
            this.hideLoadingOverlay();
          });
          // draw 이벤트가 발생하지 않을 경우를 대비한 fallback (1초 후)
          setTimeout(() => {
            const overlay = document.getElementById('tableLoadingOverlay');
            if (overlay && overlay.style.display !== 'none') {
              logger.debug('[ProjectListApp] Fallback: 오버레이 강제 숨김');
              this.hideLoadingOverlay();
            }
          }, 1000);
        } else {
          this.hideLoadingOverlay();
        }
      } else {
        this.showErrorMessage('데이터를 불러올 수 없습니다.');
        this.hideLoadingOverlay();
      }

    } catch (error) {
      logger.error('[ProjectListApp] 데이터 로딩 실패:', error);
      this.showErrorMessage('데이터를 불러오는데 실패했습니다.');
      // 에러 발생 시에도 로딩 오버레이 숨김
      this.hideLoadingOverlay();
    }
  }

  /**
   * 통계 데이터 별도 로드 (MetadataManager 사용)
   * 전문가 리뷰: "상세 데이터는 API로만 받도록"
   */
  async loadStatisticsData() {
    try {
      logger.debug('[ProjectListApp] 통계 데이터 로드 시작');

      // MetadataManager를 통한 통계 데이터 로드
      const statistics = await this.metadataManager.loadProjectStatistics();

      if (statistics) {
        // 통계 UI 업데이트
        this.metadataManager.updateStatisticsUI(statistics);
        logger.debug('[ProjectListApp] 통계 데이터 로드 완료:', statistics);
      }

    } catch (error) {
      logger.error('[ProjectListApp] 통계 데이터 로딩 실패:', error);
      // 통계 데이터는 필수가 아니므로 에러 메시지 표시하지 않음
    }
  }


  /**
   * 데이터 새로고침 (DataManager 사용)
   */
  async refreshData(forceRefresh = false, showOverlay = true, showMessage = false) {
    // 중복 호출 방지를 위한 debouncing
    if (this.isRefreshing) {
      logger.debug('[ProjectListApp] 새로고침이 이미 진행 중입니다.');
      return;
    }

    // 편집 모드 중에는 자동 새로고침 건너뛰기 (사용자가 명시적으로 새로고침 버튼을 누른 경우는 제외)
    const modeManager = window.projectListApp?.components?.accordion?.modeManager;
    if (modeManager && modeManager.isEditMode && modeManager.isEditMode() && !forceRefresh) {
      logger.debug('[ProjectListApp] 편집 모드 중이므로 자동 새로고침을 건너뜁니다.');
      return;
    }

    // 사용자가 명시적으로 새로고침 버튼을 누른 경우 (forceRefresh=true)
    // 편집 중이면 확인 요청
    if (modeManager && modeManager.isEditMode && modeManager.isEditMode() && forceRefresh) {
      const confirmRefresh = confirm('편집 중인 내용이 있습니다. 새로고침하면 저장되지 않은 변경사항이 사라집니다. 계속하시겠습니까?');
      if (!confirmRefresh) {
        logger.debug('[ProjectListApp] 사용자가 편집 중 새로고침을 취소했습니다.');
        return;
      }
      // 확인했으면 편집 모드 강제 종료 및 잠금 해제
      const accordion = window.projectListApp?.components?.accordion;
      if (accordion && accordion.currentProject) {
        const projectCode = accordion.currentProject['프로젝트 코드'];
        await accordion.disableUnifiedEditMode(projectCode, true); // skipDataRestore=true
      }
    }

    this.isRefreshing = true;
    const refreshBtn = document.getElementById('refreshProjectBtn');

    try {
      if (showMessage && refreshBtn) {
        buttonStateManager.setLoading(refreshBtn, '새로고침 중...');
      }

      // 강제 새로고침 시 테이블 로딩 오버레이 표시 (서버에서 데이터 가져올 때)
      // showOverlay가 false이면 오버레이를 표시하지 않음 (스켈레톤이 있는 경우)
      if (forceRefresh && showOverlay) {
        this.showLoadingOverlay();
      }

      // DataManager를 통한 데이터 새로고침
      const result = await this.dataManager.loadProjectData(forceRefresh);

      if (result.data) {
        // StateManager를 통한 상태 업데이트 (필터 재적용 건너뜀 - updateUI에서 처리)
        this.stateManager.setCurrentData(result.data, true);
        await this.updateUI(result.data);

        if (showMessage && refreshBtn) {
          buttonStateManager.setSuccess(refreshBtn, '새로고침');
          this.showSuccessMessage('데이터가 성공적으로 업데이트되었습니다.');

          // 2초 후 버튼을 원래 상태로 복원
          setTimeout(() => {
            buttonStateManager.reset(refreshBtn);
          }, 2000);
        }

        logger.debug(`[ProjectListApp] 데이터 새로고침 완료: ${result.data.length}개 항목`);

        // 강제 새로고침 시 DataTable draw 이벤트 리스닝
        if (forceRefresh) {
          const table = this.components.table?.table;
          if (table) {
            table.one('draw.dt', () => {
              logger.debug('[ProjectListApp] DataTable 렌더링 완료 (새로고침)');
              this.hideLoadingOverlay();
            });
            // draw 이벤트가 발생하지 않을 경우를 대비한 fallback (1초 후)
            setTimeout(() => {
              const overlay = document.getElementById('tableLoadingOverlay');
              if (overlay && overlay.style.display !== 'none') {
                logger.debug('[ProjectListApp] Fallback: 오버레이 강제 숨김 (새로고침)');
                this.hideLoadingOverlay();
              }
            }, 1000);
          } else {
            this.hideLoadingOverlay();
          }
        }
      } else if (forceRefresh) {
        this.hideLoadingOverlay();
      }

    } catch (error) {
      logger.error('[ProjectListApp] 새로고침 실패:', error);
      if (showMessage && refreshBtn) {
        buttonStateManager.setError(refreshBtn, '실패');
        this.showErrorMessage('데이터 새로고침에 실패했습니다.');
      }
      // 에러 시에도 로딩 오버레이 숨김
      if (forceRefresh) {
        this.hideLoadingOverlay();
      }
    } finally {
      // 새로고침 완료 후 플래그 해제
      this.isRefreshing = false;
    }
  }

  /**
   * UI 업데이트 (StateManager 기반)
   */
  async updateUI(data) {
    // StateManager가 이미 상태를 관리하므로 여기서는 UI 컴포넌트 업데이트만

    // 테이블 초기화 여부 확인 (재초기화하지 않음 - 한글 설정 유지)
    if (this.components.table && !this.components.table.table) {
      // 최초 초기화만 수행 (이미 초기화된 경우 건너뜀)
      if (!$.fn.dataTable.isDataTable('#projectsTable')) {
        await this.components.table.init();
      } else {
        // 이미 초기화된 인스턴스를 참조
        this.components.table.table = $('#projectsTable').DataTable();
      }
    }

    // 현대화된 필터에 데이터 전달하여 동적 옵션 생성
    if (this.components.filters && data) {
      this.components.filters.currentData = data;
      this.components.filters.isDataLoaded = true;

      // DOM 업데이트를 위한 Promise 체인으로 처리
      await new Promise(resolve => {
        // 필터 옵션 재생성
        this.components.filters.populateAllFilters(data);

        // DOM 렌더링 완료를 위한 다음 프레임 대기
        requestAnimationFrame(() => {
          // 추가 브라우저 렌더링 시간 확보
          setTimeout(() => {
            // 필터 적용 및 콜백 실행
            this.components.filters.applyFilters();
            resolve();
          }, 10);
        });
      });
    }

    // 필터된 데이터로 UI 업데이트 (StateManager의 콜백을 통해 처리됨)
    const filteredData = this.stateManager.filteredData;
    this.components.mobileView?.updateCards(filteredData);

    // 마지막 업데이트 시간 표시
    this.updateLastRefreshTime();

    logger.debug('[ProjectListApp] UI 업데이트 완료');
  }


  /**
   * 웹소켓 연결 설정
   */
  setupWebSocketConnection() {
    // 이미 연결되어 있으면 중복 연결 방지
    if (this.socket) {
      logger.debug('[Socket.IO] 이미 연결되어 있습니다. 중복 연결 방지.');
      return;
    }

    if (window.io) {
      // Socket.IO Polling 전용 설정 (개발 환경 Werkzeug 호환성)
      // WebSocket 업그레이드 시도를 막아 "Invalid frame header" 오류 방지
      const socket = window.io({
        transports: ['polling'],  // WebSocket 비활성화
        upgrade: false,           // polling에서 WebSocket으로 업그레이드 시도 방지
        timeout: 20000,           // 20초 타임아웃
        forceNew: false           // 기존 연결 재사용
      });

      // 임시로 자동 새로고침 비활성화 (무한 루프 방지)
      // socket.on('data_updated', (data) => {
      //   // 실시간 업데이트 시에는 모든 캐시를 우회하고 최신 데이터 강제 로드
      //   this.clearLocalCache();
      //   this.refreshData(true);
      // });

      // [ROCKET] 새 프로젝트 추가 이벤트 처리
      socket.on('new_project_added', (data) => {
        this.refreshData(true, false); // 오버레이 없이 새로고침 (스켈레톤 행 유지)
        this.showSuccessMessage('새 프로젝트가 추가되었습니다.');
      });

      // 공사 취소 이벤트 처리 — 자기 자신이 발생시킨 이벤트는 무시
      // (2026-07-07): 로컬 cancelConstruction이 이미 아코디언 UI + StateManager 갱신했으므로
      // 여기서 다시 stateManager.updateSingleProject를 호출하면 applyCurrentFilters가 리스트
      // 리렌더 → 아코디언 닫힘 + 스크롤 top. 다른 매니저 PC에선 정상 반영이 필요하므로 skip은
      // 자기 자신에게만.
      socket.on('project_cancelled', (data) => {
        logger.debug('[Socket.IO] 공사 취소 이벤트 수신:', data);
        if (data.sender_email && data.sender_email === window.userEmail) {
          logger.debug('[Socket.IO] 자기 자신의 취소 액션 — 재렌더 skip');
          return;
        }
        if (this.stateManager && data.updated_project) {
          this.stateManager.updateSingleProject(data.project_code, data.updated_project);
        }
        this.showSuccessMessage(`프로젝트 ${data.project_code}이(가) 취소되었습니다.`);
      });

      // 공사 재개 이벤트 처리 — 취소와 동일한 self-echo skip 규칙
      socket.on('project_resumed', (data) => {
        logger.debug('[Socket.IO] 공사 재개 이벤트 수신:', data);
        if (data.sender_email && data.sender_email === window.userEmail) {
          logger.debug('[Socket.IO] 자기 자신의 재개 액션 — 재렌더 skip');
          return;
        }
        if (this.stateManager && data.updated_project) {
          this.stateManager.updateSingleProject(data.project_code, data.updated_project);
        }
        this.showSuccessMessage(`프로젝트 ${data.project_code}이(가) 재개되었습니다.`);
      });

      socket.on('connect', () => {
        logger.debug('[Socket.IO] 소켓 연결 성공');
      });

      socket.on('disconnect', (reason) => {
        logger.debug('[Socket.IO] 소켓 연결 끊어짐:', reason);
      });

      // WebSocket 연결 에러 핸들링 (Ctrl+F5 등)
      socket.on('connect_error', (error) => {
        logger.warn('[Socket.IO] 연결 오류 (무시 가능):', error.message);
        // 에러를 무시하고 polling으로 자동 전환됨
      });

      socket.on('error', (error) => {
        logger.warn('[Socket.IO] Socket 오류:', error);
      });

      // 소켓 인스턴스를 컴포넌트들이 사용할 수 있도록 저장
      this.socket = socket;
    }
  }



  /**
   * 알림 표시 유틸리티
   *
   * 두 스코프 지원:
   *  - showSystemAlert : 사이트 최상단 헤더 (#systemAlertContainer)
   *    시스템 전역 상태 알림용 (자동 새로고침 · 앱 초기화 · 파이프라인 실패 · 네트워크 · 인증 만료)
   *  - showPageAlert   : 페이지 서브 헤더 (#headerAlertContainer)
   *    페이지 액션 결과 알림용 (프로젝트 추가·취소·재개 · 아코디언 열기 · 리드 CRUD)
   *
   * 레거시 wrapper (showSuccessMessage / showErrorMessage / showToast) 는 유지하되
   * 각 캐시점에서 명시적으로 두 스코프 중 하나를 호출하도록 리팩터됐음.
   * 기본 라우팅: success → page, error → system, generic toast → system.
   */
  showSystemAlert(message, type = 'info') {
    import('../components/Toast.js').then(({ default: Toast }) => {
      new Toast().show(message, type);
    });
    announceToScreenReader(message, type === 'error' ? 'assertive' : 'polite');
  }

  showPageAlert(message, type = 'info') {
    const container = document.getElementById('headerAlertContainer');
    if (!container) {
      // 페이지 서브 헤더 슬롯이 없으면 시스템 스코프로 fallback
      this.showSystemAlert(message, type);
      return;
    }
    const typeMap = {
      success: 'alert-success',
      error:   'alert-danger',
      danger:  'alert-danger',
      warning: 'alert-warning',
      info:    'alert-info',
    };
    const klass = typeMap[type] || 'alert-info';
    const escaped = String(message)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
    // 얇은(py-1) 알림 — alert-dismissible 의 절대위치 X 대신 flex 인라인으로 수직 중앙 정렬
    container.innerHTML = `
      <div class="alert ${klass} fade show mb-0 py-1 px-3 d-flex align-items-center" role="alert" style="font-size: 0.9rem; gap: 0.5rem;">
        <span style="flex:1;">${escaped}</span>
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="닫기" style="position:static; padding:0.35rem; margin:0; font-size:0.6rem;"></button>
      </div>
    `;
    setTimeout(() => { container.innerHTML = ''; }, 3000);
    announceToScreenReader(message, type === 'error' ? 'assertive' : 'polite');
  }

  // 레거시 wrapper — success는 페이지 성격이 더 강함 (액션 결과)
  showSuccessMessage(message) {
    this.showPageAlert(message, 'success');
  }
  // 레거시 wrapper — error는 시스템 성격이 더 강함 (앱·파이프라인)
  showErrorMessage(message) {
    this.showSystemAlert(message, 'error');
  }
  // 레거시 wrapper — 일반 toast는 시스템 (자동 새로고침 등)
  showToast(message, type = 'info') {
    this.showSystemAlert(message, type);
  }

  /**
   * 접근성(A11y) 개선 적용
   */
  enhanceAccessibility() {
    logger.debug('[A11y] 접근성 개선 적용 시작');

    // 1. Skip Link 추가 (메인 콘텐츠로 건너뛰기)
    addSkipLink();

    // 2. DataTables 접근성 개선
    const projectTable = document.querySelector('#projects-table');
    if (projectTable) {
      enhanceTableAccessibility(projectTable);
    }

    // 3. 스크린 리더 전용 스타일 추가
    if (!document.getElementById('a11y-styles')) {
      const style = document.createElement('style');
      style.id = 'a11y-styles';
      style.textContent = `
        .sr-only {
          position: absolute;
          width: 1px;
          height: 1px;
          padding: 0;
          margin: -1px;
          overflow: hidden;
          clip: rect(0, 0, 0, 0);
          white-space: nowrap;
          border-width: 0;
        }
        .sr-only-focusable:focus {
          position: static;
          width: auto;
          height: auto;
          overflow: visible;
          clip: auto;
          white-space: normal;
        }
      `;
      document.head.appendChild(style);
    }

    logger.debug('[A11y] 접근성 개선 적용 완료');
  }

  /**
   * 마지막 업데이트 시간 표시
   */
  updateLastRefreshTime() {
    const timeElement = document.getElementById('lastRefreshTime');
    if (timeElement) {
      const now = new Date();
      timeElement.textContent = `마지막 업데이트: ${now.toLocaleTimeString('ko-KR')}`;
    }
  }


  /**
   * 앱 정리 (페이지 언로드 시) - 분리된 서비스들 정리
   */
  cleanup() {
    logger.debug('[ProjectListApp] 정리 작업 시작');

    // Socket 연결 먼저 정리 (중요!)
    if (this.socket) {
      try {
        logger.debug('[ProjectListApp] Socket 연결 해제 중...');
        // 이미 끊어진 소켓이면 무시
        if (this.socket.connected) {
          this.socket.removeAllListeners();  // 모든 이벤트 리스너 제거
          this.socket.disconnect();
        }
        this.socket = null;
        logger.debug('[ProjectListApp] Socket 연결 해제 완료');
      } catch (error) {
        logger.warn('[ProjectListApp] Socket 해제 실패:', error);
      }
    }

    // 유틸리티 정리
    try {
      buttonStateManager.cleanup();
    } catch (error) {
      logger.warn('[ProjectListApp] buttonStateManager 정리 실패:', error);
    }

    // 서비스들 정리
    if (this.stateManager) {
      try {
        this.stateManager.cleanup();
      } catch (error) {
        logger.warn('[ProjectListApp] stateManager 정리 실패:', error);
      }
    }

    if (this.metadataManager) {
      try {
        this.metadataManager.cleanup();
      } catch (error) {
        logger.warn('[ProjectListApp] metadataManager 정리 실패:', error);
      }
    }

    if (this.cacheStatusManager) {
      try {
        this.cacheStatusManager.cleanup();
      } catch (error) {
        logger.warn('[ProjectListApp] cacheStatusManager 정리 실패:', error);
      }
    }

    // 컴포넌트들 정리
    Object.values(this.components).forEach(component => {
      if (component && typeof component.cleanup === 'function') {
        try {
          component.cleanup();
        } catch (error) {
          logger.warn('[ProjectListApp] 컴포넌트 정리 실패:', error);
        }
      }
    });

    // 초기화 플래그 리셋 (재초기화 가능하도록)
    this.isInitialized = false;
    logger.debug('[ProjectListApp] isInitialized 플래그 리셋');

    logger.debug('[ProjectListApp] 정리 작업 완료');
  }
}

// 싱글톤 패턴: 전역에 이미 인스턴스가 있으면 재사용
let app;

// 모듈 중복 로드 방지
if (window.__projectListAppModuleLoaded) {
  logger.warn('[ProjectListApp] 모듈이 이미 로드되었습니다. 중복 실행 방지');
  app = window.__projectListAppInstance;
} else {
  logger.debug('[ProjectListApp] 모듈 최초 로드');
  window.__projectListAppModuleLoaded = true;

  // 이전 인스턴스가 있으면 강제 cleanup (빠른 새로고침 대응)
  if (window.__projectListAppInstance) {
    logger.debug('[ProjectListApp] 이전 인스턴스 정리 중...');
    try {
      window.__projectListAppInstance.cleanup();
    } catch (error) {
      logger.warn('[ProjectListApp] 이전 인스턴스 정리 실패:', error);
    }
    window.__projectListAppInstance = null;
  }

  logger.debug('[ProjectListApp] 새 인스턴스 생성');
  app = new ProjectListApp();
  window.__projectListAppInstance = app;
  window.ProjectListApp = app; // 레거시 호환성
  window.projectListApp = app; // DataSyncManager 호환성 (소문자)

  // 알림 헬퍼 window 전역 노출 — 다른 컴포넌트(ProjectRowAccordion·ModernProjectModal 등)가
  // app 인스턴스 참조 없이 직접 호출할 수 있게. showSystemAlert = 사이트 최상단 헤더,
  // showPageAlert = 페이지 서브 헤더. 스코프 분류 규칙은 app 클래스 정의부 주석 참조.
  window.showSystemAlert = (msg, type) => app.showSystemAlert(msg, type);
  window.showPageAlert = (msg, type) => app.showPageAlert(msg, type);

  // DOM 준비 완료 시 초기화
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => app.init());
  } else {
    app.init();
  }

  // 페이지 언로드 시 정리 (beforeunload + pagehide 둘 다 처리)
  const handlePageUnload = () => {
    logger.debug('[ProjectListApp] 페이지 언로드 - 정리 시작');

    // 프로젝트 잠금 해제 (동기 방식으로 즉시 전송)
    try {
      // sendBeacon을 사용하여 페이지 종료 시에도 요청이 전송되도록 함
      const lockReleaseData = JSON.stringify({});
      const lockReleaseUrl = '/api/project-lock/release-all-user-locks';

      if (navigator.sendBeacon) {
        // sendBeacon 사용 (페이지 종료 시에도 전송 보장)
        const blob = new Blob([lockReleaseData], { type: 'application/json' });
        navigator.sendBeacon(lockReleaseUrl, blob);
        logger.debug('[ProjectListApp] 프로젝트 잠금 해제 요청 전송 (sendBeacon)');
      } else {
        // sendBeacon 미지원 시 동기 fetch 사용
        fetch(lockReleaseUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: lockReleaseData,
          keepalive: true  // 페이지 종료 후에도 요청 유지
        }).catch(err => {
          logger.warn('[ProjectListApp] 잠금 해제 요청 실패 (무시됨):', err);
        });
        logger.debug('[ProjectListApp] 프로젝트 잠금 해제 요청 전송 (fetch)');
      }
    } catch (error) {
      logger.warn('[ProjectListApp] 프로젝트 잠금 해제 실패:', error);
    }

    app.cleanup();
    window.__projectListAppModuleLoaded = false; // 플래그 리셋
  };

  window.addEventListener('beforeunload', handlePageUnload);
  window.addEventListener('pagehide', handlePageUnload);  // iOS Safari 및 Ctrl+F5 대응

  // 페이지가 다시 보일 때 5분 이상 경과 시에만 새로고침 (세밀한 상태 제어)
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && app.isInitialized) {
      const now = Date.now();
      const timeSinceLastChange = now - app.lastVisibilityChange;

      // 5분(300초) 이상 떨어져 있었을 때만 새로고침
      const REFRESH_THRESHOLD = 5 * 60 * 1000; // 5분

      if (timeSinceLastChange < REFRESH_THRESHOLD) {
        logger.debug(`🔒 [새로고침 차단] 마지막 변경 후 ${Math.round(timeSinceLastChange / 1000)}초만 지남 (최소 ${REFRESH_THRESHOLD / 1000}초 필요)`);
        app.lastVisibilityChange = now;
        return;
      }

      // 편집 모드 체크 (실제 편집 상태 확인)
      const isEditingMode = document.querySelector('.accordion-row.open .editing') || // 아코디언 내부 편집 중인 카드
                            document.querySelector('.accordion-content .editing') ||   // 아코디언 내 편집 요소
                            document.querySelector('input:focus, textarea:focus, select:focus'); // 포커스된 입력 필드

      if (isEditingMode) {
        logger.debug('🔒 [새로고침 차단] 편집 중이라 자동 새로고침 건너뜀');
        app.lastVisibilityChange = now;
        return;
      }

      // 아코디언이 열려있으면 새로고침 건너뛰기
      const openAccordion = document.querySelector('.accordion-row.open');
      if (openAccordion) {
        logger.debug('🔒 [새로고침 차단] 아코디언이 열려있어 자동 새로고침 건너뜀');
        app.lastVisibilityChange = now;
        return;
      }

      logger.debug('🔄 [자동 새로고침] 5분 이상 경과, 데이터 새로고침 시작');
      app.lastVisibilityChange = now;

      // 자동 새로고침 (오버레이 없이)
      app.refreshData(false, false);

      // 수금 관리 모드에서는 토스트를 덜 방해되게 표시
      const isCollectionMode = app.components?.filters?.filters?.outstanding === 'outstanding';

      if (isCollectionMode) {
        // 수금 관리 모드에서는 간단한 메시지만 로그에 기록
        logger.debug('ℹ️ [수금 모드] 자동 갱신 완료 (토스트 생략)');
      } else {
        // 일반 모드에서는 토스트 표시
        app.showToast('최신 데이터로 갱신되었습니다', 'info');
      }
    }
  });
}

// 모던화된 API 노출
const projectAPI = exposeGlobalAPI(app);

// DataSyncManager: 프로젝트 생성 후 실시간 테이블 업데이트
window.DataSyncManager = {
  /**
   * DataTable 초기화 대기 (최대 10초)
   */
  async waitForDataTable(maxWaitTime = 10000) {
    const startTime = Date.now();
    let loggedOnce = false;

    while (true) {
      const elapsed = Date.now() - startTime;

      // ProjectTable 컴포넌트의 table 인스턴스가 있는지 확인 (더 정확한 방법)
      const isTableReady = window.projectListApp?.components?.table?.table != null;

      if (isTableReady) {
        logger.debug('[DataSyncManager] DataTable 초기화 완료 확인 (ProjectTable.table 인스턴스 존재)');
        return true;
      }

      // 2초 후에 한 번만 로그 출력
      if (!loggedOnce && elapsed > 2000) {
        logger.debug('[DataSyncManager] DataTable 초기화 대기 중... (ProjectTable 컴포넌트 로딩 중)');
        loggedOnce = true;
      }

      if (elapsed > maxWaitTime) {
        logger.warn('[DataSyncManager] DataTable 초기화 대기 시간 초과 (10초)');
        logger.warn('[DataSyncManager] 페이지 로드 직후라면 정상입니다. 데이터 로딩 완료 후 새로고침으로 프로젝트가 표시됩니다.');
        return false;
      }

      await new Promise(resolve => setTimeout(resolve, 100));
    }
  },

  addProjectRealTime(projectData) {
    logger.debug('[DataSyncManager] 새 프로젝트 추가 시작');
    logger.debug('[DataSyncManager] 받은 데이터:', JSON.stringify(projectData, null, 2));

    // ProjectTable 컴포넌트의 table 인스턴스 확인
    const table = window.projectListApp?.components?.table?.table;
    if (!table) {
      logger.debug('[DataSyncManager] DataTable 초기화 전이므로 데이터 새로고침으로 처리');
      // 오버레이 없이 조용히 새로고침 (프로젝트가 이미 저장되었으므로)
      this.refreshData(false);
      return;
    }

    try {
      // 데이터 무결성 검증: 필수 필드가 모두 있는지 확인
      const requiredFields = ['프로젝트 코드', '담당자', '유입 구분'];
      const hasAllFields = requiredFields.every(field =>
        projectData.hasOwnProperty(field)
      );

      if (!hasAllFields) {
        logger.warn('[DataSyncManager] 필수 필드가 누락되어 전체 새로고침 실행:', projectData);
        this.refreshData(true);
        return;
      }

      // localStorage 캐시 무효화 (다음 새로고침 시 서버에서 최신 데이터 가져오기)
      if (window.projectListApp?.dataManager?.clearCache) {
        window.projectListApp.dataManager.clearCache();
        logger.debug('[DataSyncManager] 캐시 무효화 완료');
      }

      // 새 행 추가 (DataTables API 사용)
      const newRow = table.row.add(projectData).draw(false);
      logger.debug('[DataSyncManager] 테이블에 새 행 추가 완료');

      // StateManager 업데이트 (리스너 자동 호출로 필터/모바일 카드 등 모든 구독자 갱신)
      const stateManager = window.projectListApp?.stateManager;
      if (stateManager) {
        const newData = [...stateManager.currentData, projectData];
        stateManager.setCurrentData(newData);
        logger.debug('[DataSyncManager] StateManager 데이터 업데이트 완료 (dataChange 리스너 통지됨)');
      }

      // 추가된 행에 스크롤 및 하이라이트
      const rowNode = newRow.node();
      if (rowNode) {
        // 부드러운 스크롤
        setTimeout(() => {
          rowNode.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 100);

        // 하이라이트 애니메이션 (확장/축소 효과)
        rowNode.classList.add('new-row-pulse');
        logger.debug('[DataSyncManager] 새 행에 확장/축소 애니메이션 적용');

        setTimeout(() => {
          rowNode.classList.remove('new-row-pulse');
        }, 2000);
      }
    } catch (error) {
      logger.error('[DataSyncManager] 행 추가 실패:', error);
      logger.warn('[DataSyncManager] 실패로 인해 전체 새로고침 실행');
      this.refreshData(true);
    }
  },

  refreshData(showOverlay = true) {
    logger.debug('[DataSyncManager] 데이터 새로고침 (force=true, showOverlay=' + showOverlay + ')');

    if (app && app.refreshData) {
      app.refreshData(true, showOverlay); // forceRefresh=true로 캐시 무시
    } else if (app && app.dataManager && app.dataManager.loadData) {
      app.dataManager.loadData(true); // force refresh
    }
  }
};

// ES6 모듈 export
export { app as ProjectListApp, projectAPI };
export default ProjectListApp;