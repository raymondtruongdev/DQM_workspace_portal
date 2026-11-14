/** Thay `@microsoft/sp-http` khi đóng gói khung chạy thử (module không dùng SPHttpClient thật). */
export class SPHttpClient {
  public static configurations = { v1: {} };
}
export type SPHttpClientResponse = Response;
