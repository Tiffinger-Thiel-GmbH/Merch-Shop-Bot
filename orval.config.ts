import { defineConfig } from "orval";

export default defineConfig({
  merchBot: {
    input: "http://localhost:3000/api-json",
    output: {
      httpClient: "axios",
      target: "./src/api/merchApi.ts",
      tsconfig: "./tsconfig.json",
      override: {
        mutator: {
          path: "./src/api/axiosInstance.ts",
          name: "customInstance",
        },
      },
    },
  },
});
