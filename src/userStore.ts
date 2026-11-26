import { createStore } from "zustand/vanilla";

interface UserState {
  userId: string | null;
  token: string | null;
  setUser: (userId: string, token: string) => void;
  clear: () => void;
}

export const useUserStore = createStore<UserState>()((set) => ({
  userId: null,
  token: null,
  setUser: (userId, token) => set({ userId, token }),
  clear: () => set({ userId: null, token: null }),
}));
