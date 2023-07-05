import {HttpClient} from  '@microsoft/sp-http'
export interface IDurableFunctionsStatusProps {
  baseUrl: string; taskHub: string; systemKey: string;
  httpClient:HttpClient;
}
