import type { Dataset } from './data';

const DB_NAME = 'cauhs-local-data';
const STORE_NAME = 'datasets';
const LAST_UPLOAD_KEY = 'last-upload';

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) {
      reject(new Error('이 브라우저에서는 로컬 저장을 지원하지 않습니다.'));
      return;
    }
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) {
        request.result.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('로컬 저장소를 열지 못했습니다.'));
    request.onblocked = () => reject(new Error('다른 탭에서 저장소를 사용 중입니다. 탭을 새로고침한 뒤 다시 시도해 주세요.'));
  });
}

export async function readLastUpload(): Promise<Dataset | null> {
  const db = await openDatabase();
  try {
    return await new Promise((resolve, reject) => {
      const request = db.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).get(LAST_UPLOAD_KEY);
      request.onsuccess = () => resolve((request.result as Dataset | undefined) ?? null);
      request.onerror = () => reject(request.error ?? new Error('저장된 엑셀 데이터를 읽지 못했습니다.'));
    });
  } finally {
    db.close();
  }
}

export async function saveLastUpload(dataset: Dataset): Promise<void> {
  const db = await openDatabase();
  try {
    await new Promise<void>((resolve, reject) => {
      const request = db.transaction(STORE_NAME, 'readwrite').objectStore(STORE_NAME).put(dataset, LAST_UPLOAD_KEY);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error ?? new Error('엑셀 데이터를 브라우저에 저장하지 못했습니다.'));
    });
  } finally {
    db.close();
  }
}

export async function clearLastUpload(): Promise<void> {
  const db = await openDatabase();
  try {
    await new Promise<void>((resolve, reject) => {
      const request = db.transaction(STORE_NAME, 'readwrite').objectStore(STORE_NAME).delete(LAST_UPLOAD_KEY);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error ?? new Error('저장된 엑셀 데이터를 삭제하지 못했습니다.'));
    });
  } finally {
    db.close();
  }
}
