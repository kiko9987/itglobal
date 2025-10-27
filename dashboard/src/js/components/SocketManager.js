import logger from '../utils/logger.js';

/**
 * SocketManager 컴포넌트
 * 실시간 데이터 동기화를 위한 중앙집중식 WebSocket 관리 시스템
 */
export default class SocketManager {
  constructor(options = {}) {
    // 설정
    this.config = {
      autoConnect: true,
      reconnectAttempts: 5,
      reconnectInterval: 1000,
      heartbeatInterval: 30000,
      maxReconnectInterval: 30000,
      timeout: 10000,
      debug: false,
      ...options
    };

    // 상태 관리
    this.socket = null;
    this.isConnected = false;
    this.isConnecting = false;
    this.connectionAttempts = 0;
    this.lastConnectedTime = null;
    this.reconnectTimer = null;
    this.heartbeatTimer = null;

    // 이벤트 구독자 관리
    this.subscribers = new Map();
    this.messageQueue = [];
    this.pendingRequests = new Map();

    // 연결 상태 리스너들
    this.connectionListeners = new Set();

    // 현재 사용자 정보
    this.currentUser = null;

    this.initializeUser();

    if (this.config.autoConnect) {
      this.connect();
    }
  }

  /**
   * 사용자 정보 초기화
   */
  initializeUser() {
    this.currentUser = {
      id: window.userRole || 'anonymous',
      name: window.session?.user?.name || '익명 사용자',
      sessionId: this.generateSessionId(),
      timestamp: Date.now()
    };
  }

  /**
   * 세션 ID 생성
   * crypto.randomUUID() 사용으로 암호학적으로 안전한 UUID 생성
   */
  generateSessionId() {
    // 브라우저가 crypto.randomUUID()를 지원하면 사용
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
      return crypto.randomUUID();
    }

    // 폴백: crypto.getRandomValues() 사용
    if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
      const array = new Uint8Array(16);
      crypto.getRandomValues(array);
      // UUID v4 형식으로 변환
      array[6] = (array[6] & 0x0f) | 0x40; // version 4
      array[8] = (array[8] & 0x3f) | 0x80; // variant
      const hex = Array.from(array, byte => byte.toString(16).padStart(2, '0')).join('');
      return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
    }

    // 최종 폴백 (매우 오래된 브라우저)
    logger.warn('[SocketManager] crypto API를 사용할 수 없습니다. 약한 난수 생성기를 사용합니다.');
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  }

  /**
   * WebSocket 연결
   */
  connect() {
    if (this.isConnected || this.isConnecting) {
      this.log('이미 연결되어 있거나 연결 중입니다.');
      return Promise.resolve();
    }

    return new Promise((resolve, reject) => {
      try {
        this.isConnecting = true;
        this.log('WebSocket 연결 시도 중...');

        // Socket.IO 연결
        if (window.io) {
          this.socket = window.io({
            transports: ['polling'],  // WebSocket 비활성화 (개발 환경 Werkzeug 호환성)
            upgrade: false,           // polling에서 WebSocket으로 업그레이드 시도 방지
            timeout: this.config.timeout,
            forceNew: true
          });

          this.setupSocketEvents(resolve, reject);
        } else {
          throw new Error('Socket.IO 라이브러리가 로드되지 않았습니다.');
        }
      } catch (error) {
        this.isConnecting = false;
        this.handleConnectionError(error);
        reject(error);
      }
    });
  }

  /**
   * Socket 이벤트 설정
   */
  setupSocketEvents(resolve, reject) {
    // 연결 성공
    this.socket.on('connect', () => {
      this.handleConnectionSuccess();
      resolve();
    });

    // 연결 실패
    this.socket.on('connect_error', (error) => {
      this.handleConnectionError(error);
      if (this.isConnecting) {
        reject(error);
      }
    });

    // 연결 끊김
    this.socket.on('disconnect', (reason) => {
      this.handleDisconnection(reason);
    });

    // 재연결 성공
    this.socket.on('reconnect', (attemptNumber) => {
      this.log(`재연결 성공 (${attemptNumber}번째 시도)`);
      this.handleReconnection();
    });

    // 재연결 시도
    this.socket.on('reconnect_attempt', (attemptNumber) => {
      this.log(`재연결 시도 중... (${attemptNumber}/${this.config.reconnectAttempts})`);
    });

    // 재연결 실패
    this.socket.on('reconnect_failed', () => {
      this.log('재연결 실패');
      this.handleReconnectionFailure();
    });

    // 사용자 정의 이벤트들
    this.setupCustomEvents();
  }

  /**
   * 사용자 정의 이벤트 설정
   */
  setupCustomEvents() {
    // 프로젝트 데이터 업데이트
    this.socket.on('project_updated', (data) => {
      this.emit('project_updated', data);
    });

    // 프로젝트 잠금 상태 변경 (새 이벤트)
    this.socket.on('project_lock_changed', (data) => {
      this.emit('project_lock_changed', data);
    });

    // 사용자 활동 알림
    this.socket.on('user_activity', (data) => {
      this.emit('user_activity', data);
    });

    // 시스템 알림
    this.socket.on('system_notification', (data) => {
      this.emit('system_notification', data);
    });

    // 데이터 동기화
    this.socket.on('data_sync', (data) => {
      this.emit('data_sync', data);
    });

    // 에러 알림
    this.socket.on('error_notification', (data) => {
      this.emit('error_notification', data);
    });

    // 하트비트 응답
    this.socket.on('heartbeat_response', (data) => {
      this.handleHeartbeatResponse(data);
    });
  }

  /**
   * 연결 성공 처리
   */
  handleConnectionSuccess() {
    this.isConnected = true;
    this.isConnecting = false;
    this.connectionAttempts = 0;
    this.lastConnectedTime = Date.now();

    this.log('[SUCCESS] WebSocket 연결 성공');

    // 사용자 등록
    this.registerUser();

    // 대기 중인 메시지 전송
    this.flushMessageQueue();

    // 하트비트 시작
    this.startHeartbeat();

    // 연결 상태 알림
    this.notifyConnectionListeners('connected');

    // 연결 성공 이벤트 발생
    this.emit('socket_connected', {
      timestamp: this.lastConnectedTime,
      user: this.currentUser
    });
  }

  /**
   * 연결 오류 처리
   */
  handleConnectionError(error) {
    this.isConnecting = false;
    this.connectionAttempts++;

    this.log(`[ERROR] 연결 실패 (${this.connectionAttempts}/${this.config.reconnectAttempts}):`, error.message);

    // 재연결 시도
    if (this.connectionAttempts < this.config.reconnectAttempts) {
      this.scheduleReconnect();
    } else {
      this.handleReconnectionFailure();
    }

    this.emit('socket_error', {
      error,
      attempts: this.connectionAttempts,
      maxAttempts: this.config.reconnectAttempts
    });
  }

  /**
   * 연결 끊김 처리
   */
  handleDisconnection(reason) {
    this.isConnected = false;
    this.lastConnectedTime = null;

    this.log(`[ERROR] 연결 끊김: ${reason}`);

    // 하트비트 중지
    this.stopHeartbeat();

    // 연결 상태 알림
    this.notifyConnectionListeners('disconnected');

    // 자동 재연결 (서버에서 끊은 경우가 아니라면)
    if (reason !== 'io server disconnect') {
      this.scheduleReconnect();
    }

    this.emit('socket_disconnected', {
      reason,
      timestamp: Date.now()
    });
  }

  /**
   * 재연결 성공 처리
   */
  handleReconnection() {
    this.handleConnectionSuccess();
    this.emit('socket_reconnected', {
      timestamp: Date.now(),
      attempts: this.connectionAttempts
    });
  }

  /**
   * 재연결 실패 처리
   */
  handleReconnectionFailure() {
    this.isConnected = false;
    this.isConnecting = false;

    this.log('[ERROR] 재연결 최대 시도 횟수 초과');

    this.notifyConnectionListeners('failed');
    this.emit('socket_reconnect_failed', {
      attempts: this.connectionAttempts,
      maxAttempts: this.config.reconnectAttempts
    });
  }

  /**
   * 재연결 스케줄링
   */
  scheduleReconnect() {
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
    }

    const delay = Math.min(
      this.config.reconnectInterval * Math.pow(2, this.connectionAttempts - 1),
      this.config.maxReconnectInterval
    );

    this.log(`${delay}ms 후 재연결 시도...`);

    this.reconnectTimer = setTimeout(() => {
      this.connect().catch(error => {
        this.log('재연결 실패:', error.message);
      });
    }, delay);
  }

  /**
   * 사용자 등록
   */
  registerUser() {
    if (this.isConnected && this.socket) {
      this.socket.emit('register_user', this.currentUser);
    }
  }

  /**
   * 하트비트 시작
   */
  startHeartbeat() {
    this.stopHeartbeat();

    this.heartbeatTimer = setInterval(() => {
      if (this.isConnected && this.socket) {
        this.socket.emit('heartbeat', {
          timestamp: Date.now(),
          user: this.currentUser
        });
      }
    }, this.config.heartbeatInterval);
  }

  /**
   * 하트비트 중지
   */
  stopHeartbeat() {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
  }

  /**
   * 하트비트 응답 처리
   */
  handleHeartbeatResponse(data) {
    // 서버 시간과 동기화, 지연 시간 계산 등
    const latency = Date.now() - data.clientTimestamp;
    this.emit('heartbeat', { latency, serverTime: data.serverTime });
  }

  /**
   * 메시지 전송
   */
  send(event, data, callback) {
    const message = {
      event,
      data,
      timestamp: Date.now(),
      id: this.generateMessageId()
    };

    if (this.isConnected && this.socket) {
      this.socket.emit(event, data, callback);
      this.log(`📤 메시지 전송: ${event}`, data);
    } else {
      // 연결되지 않은 경우 큐에 저장
      this.messageQueue.push({ message, callback });
      this.log(`📦 메시지 큐에 저장: ${event}`, data);
    }

    return message.id;
  }

  /**
   * 메시지 ID 생성
   */
  generateMessageId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  }

  /**
   * 대기 중인 메시지 큐 전송
   */
  flushMessageQueue() {
    if (!this.isConnected || !this.socket) return;

    this.log(`📤 ${this.messageQueue.length}개의 대기 메시지 전송`);

    const queue = [...this.messageQueue];
    this.messageQueue = [];

    queue.forEach(({ message, callback }) => {
      this.socket.emit(message.event, message.data, callback);
    });
  }

  /**
   * 이벤트 구독
   */
  on(event, handler) {
    if (!this.subscribers.has(event)) {
      this.subscribers.set(event, new Set());
    }
    this.subscribers.get(event).add(handler);

    this.log(`📝 이벤트 구독: ${event}`);
  }

  /**
   * 이벤트 구독 해제
   */
  off(event, handler) {
    if (this.subscribers.has(event)) {
      this.subscribers.get(event).delete(handler);
      if (this.subscribers.get(event).size === 0) {
        this.subscribers.delete(event);
      }
    }
  }

  /**
   * 이벤트 발생
   */
  emit(event, data) {
    if (this.subscribers.has(event)) {
      this.subscribers.get(event).forEach(handler => {
        try {
          handler(data);
        } catch (error) {
          this.log(`이벤트 핸들러 오류 (${event}):`, error);
        }
      });
    }
  }

  /**
   * 일회성 이벤트 구독
   */
  once(event, handler) {
    const onceHandler = (data) => {
      handler(data);
      this.off(event, onceHandler);
    };
    this.on(event, onceHandler);
  }

  /**
   * 연결 상태 리스너 추가
   */
  onConnectionChange(listener) {
    this.connectionListeners.add(listener);
  }

  /**
   * 연결 상태 리스너 제거
   */
  offConnectionChange(listener) {
    this.connectionListeners.delete(listener);
  }

  /**
   * 연결 상태 알림
   */
  notifyConnectionListeners(status) {
    this.connectionListeners.forEach(listener => {
      try {
        listener(status, {
          isConnected: this.isConnected,
          connectionAttempts: this.connectionAttempts,
          lastConnectedTime: this.lastConnectedTime
        });
      } catch (error) {
        this.log('연결 상태 리스너 오류:', error);
      }
    });
  }

  /**
   * 수동 연결 해제
   */
  disconnect() {
    this.log('수동 연결 해제');

    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }

    this.stopHeartbeat();

    if (this.socket) {
      this.socket.disconnect();
      this.socket = null;
    }

    this.isConnected = false;
    this.isConnecting = false;
    this.connectionAttempts = 0;

    this.notifyConnectionListeners('disconnected');
  }

  /**
   * 수동 재연결
   */
  reconnect() {
    this.disconnect();
    return this.connect();
  }

  /**
   * 연결 상태 조회
   */
  getConnectionStatus() {
    return {
      isConnected: this.isConnected,
      isConnecting: this.isConnecting,
      connectionAttempts: this.connectionAttempts,
      lastConnectedTime: this.lastConnectedTime,
      messageQueueLength: this.messageQueue.length,
      subscriberCount: this.subscribers.size
    };
  }

  /**
   * 디버그 로그
   */
  log(...args) {
    if (this.config.debug) {
      logger.debug('[SocketManager]', ...args);
    }
  }

  /**
   * 특정 사용자에게 메시지 전송
   */
  sendToUser(userId, event, data) {
    this.send('user_message', {
      targetUserId: userId,
      event,
      data
    });
  }

  /**
   * 브로드캐스트 메시지 전송
   */
  broadcast(event, data) {
    this.send('broadcast', {
      event,
      data
    });
  }

  /**
   * 룸 조인
   */
  joinRoom(roomName) {
    this.send('join_room', { room: roomName });
  }

  /**
   * 룸 나가기
   */
  leaveRoom(roomName) {
    this.send('leave_room', { room: roomName });
  }

  /**
   * 룸에 메시지 전송
   */
  sendToRoom(roomName, event, data) {
    this.send('room_message', {
      room: roomName,
      event,
      data
    });
  }

  /**
   * 동기화 요청
   */
  requestSync(dataType) {
    this.send('request_sync', { dataType });
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    this.log('SocketManager 정리 중...');

    this.disconnect();
    this.subscribers.clear();
    this.connectionListeners.clear();
    this.messageQueue = [];
    this.pendingRequests.clear();

    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
    }
  }
}

// 전역 SocketManager 인스턴스 (싱글톤 패턴)
let globalSocketManager = null;

export function getSocketManager(options = {}) {
  if (!globalSocketManager) {
    globalSocketManager = new SocketManager(options);
  }
  return globalSocketManager;
}

export function destroySocketManager() {
  if (globalSocketManager) {
    globalSocketManager.destroy();
    globalSocketManager = null;
  }
}