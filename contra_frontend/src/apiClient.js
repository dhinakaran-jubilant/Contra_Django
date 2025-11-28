// apiClient.js
import axios from "axios";

const apiClient = axios.create({
  baseURL: "http://127.0.0.1:8000/api/v1/",
  // baseURL: "http://192.168.0.7:8000/api/v1/",
  withCredentials: true,         
});

export default apiClient;
