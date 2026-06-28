/**
 * Simple, robust IndexedDB key-value store for storing large datasets (restored databases, readings)
 * to avoid localStorage's 5MB quota limits.
 */

const DB_NAME = 'energy-monitor-db';
const STORE_NAME = 'keyvalue';
const DB_VERSION = 1;

function getDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = () => {
      reject(request.error);
    };

    request.onsuccess = () => {
      resolve(request.result);
    };

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
  });
}

export function dbGet<T>(key: string): Promise<T | null> {
  return getDB().then((db) => {
    return new Promise<T | null>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(key);

      request.onsuccess = () => {
        resolve(request.result !== undefined ? request.result as T : null);
      };

      request.onerror = () => {
        reject(request.error);
      };
    });
  }).catch((err) => {
    console.error(`IndexedDB dbGet failed for key "${key}", falling back to localStorage:`, err);
    try {
      const val = localStorage.getItem(key);
      return val ? JSON.parse(val) as T : null;
    } catch {
      return null;
    }
  });
}

export function dbSet<T>(key: string, value: T): Promise<void> {
  return getDB().then((db) => {
    return new Promise<void>((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.put(value, key);

      request.onsuccess = () => {
        resolve();
      };

      request.onerror = () => {
        reject(request.error);
      };
    });
  }).catch((err) => {
    console.error(`IndexedDB dbSet failed for key "${key}", falling back to localStorage:`, err);
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (e) {
      console.error('LocalStorage write fallback also failed:', e);
    }
  });
}
