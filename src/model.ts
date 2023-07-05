export interface IInstance{
    name:string;
    instanceId:string;
    runtimeStatus:string;
    input:string;
    customStatus:string;
    output:string;
    createdTime:Date;
    lastUpdatedTime:Date;
    historyEvents?:Array<IHistoryEvent>
}
export interface IHistoryEvent{
    Correlation:string;
    EventType:string;
    FunctionName:string;
    Generation:number;
    Input:string;
    Result:string;
    ScheduledTime:Date;
  }