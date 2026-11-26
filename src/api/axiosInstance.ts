import Axios, { AxiosRequestConfig, AxiosError } from "axios";
import { useUserStore } from "../userStore";

export const AXIOS_INSTANCE = Axios.create({
  baseURL: process.env.API_BASE_URL || "http://localhost:3000", // use your own URL or environment variable
});

AXIOS_INSTANCE.interceptors.request.use((req) => {
  const { userId } = useUserStore.getState();
  console.log(`Attaching x-user-id ${userId},`);
  if (userId) {
    req.headers["x-user-id"] = userId;
  }
  return req;
});

// Add a second `options` argument to pass extra options to each query
export const customInstance = <T>(
  config: AxiosRequestConfig,
  options?: AxiosRequestConfig,
): Promise<T> => {
  const promise = AXIOS_INSTANCE({
    ...config,
    ...options,
  }).then(({ data }) => data);

  return promise;
};

// Override the return error type for react-query and swr
export type ErrorType<Error> = AxiosError<Error>;

// Standard body type
export type BodyType<BodyData> = BodyData;

// Or wrap the body type if processing data before sending
// export type BodyType<BodyData> = CamelCase<BodyData>;
