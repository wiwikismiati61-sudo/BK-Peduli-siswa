
const DB_NAME = "BK_Peduli_DB_v3";
const DB_VERSION = 1;

export class DBService {
  private db: IDBDatabase | null = null;

  async init(): Promise<void> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (e: any) => {
        const d = e.target.result;
        if (!d.objectStoreNames.contains("kasus")) d.createObjectStore("kasus", { keyPath: "id", autoIncrement: true });
        if (!d.objectStoreNames.contains("siswa")) d.createObjectStore("siswa", { keyPath: "id", autoIncrement: true });
        if (!d.objectStoreNames.contains("wali_kelas")) d.createObjectStore("wali_kelas", { keyPath: "id", autoIncrement: true });
        if (!d.objectStoreNames.contains("guru_bk")) d.createObjectStore("guru_bk", { keyPath: "id", autoIncrement: true });
      };

      request.onsuccess = (e: any) => {
        this.db = e.target.result;
        resolve();
      };

      request.onerror = (e) => reject(e);
    });
  }

  async getAll<T>(storeName: string): Promise<T[]> {
    if (!this.db) await this.init();
    return new Promise((resolve) => {
      const tx = this.db!.transaction(storeName, "readonly");
      const store = tx.objectStore(storeName);
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve([]);
    });
  }

  async put(storeName: string, data: any): Promise<void> {
    if (!this.db) await this.init();
    return new Promise((resolve) => {
      const tx = this.db!.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      store.put(data);
      tx.oncomplete = () => resolve();
    });
  }

  async delete(storeName: string, id: number): Promise<void> {
    if (!this.db) await this.init();
    return new Promise((resolve) => {
      const tx = this.db!.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      store.delete(id);
      tx.oncomplete = () => resolve();
    });
  }

  async clear(storeName: string): Promise<void> {
    if (!this.db) await this.init();
    return new Promise((resolve) => {
      const tx = this.db!.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      store.clear();
      tx.oncomplete = () => resolve();
    });
  }
}

export const dbService = new DBService();
