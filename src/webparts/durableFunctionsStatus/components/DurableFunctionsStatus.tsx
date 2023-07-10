import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { IHistoryEvent, IInstance } from '../../../model'
import styles from './DurableFunctionsStatus.module.scss';
import { IDurableFunctionsStatusProps } from './IDurableFunctionsStatusProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient } from '@microsoft/sp-http'
import { orderBy, result, sortBy } from 'lodash';
import { Button, DetailsList, Fabric, IColumn, IDetailsRowProps, Link, PrimaryButton, TextField, DetailsRow, CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react';
import { render } from 'react-dom';

import { format, formatDuration, intervalToDuration } from 'date-fns';
import { utcToZonedTime } from 'date-fns-tz';

export default function DurableFunctionsStatus(props: IDurableFunctionsStatusProps): JSX.Element {
  const [selectedInstance, setSelectedInstance] = React.useState<IInstance>(null);
  const [instances, setInstances] = useState<Array<IInstance>>([])
  const [refreshSeconds, setRefreshSeconds] = useState<number>(null);
  const [refreshDescription, setRefreshDescription] = useState<string>(null);
  const instanceIntervalRef = useRef<number | null>(null);
  //* See https://www.kindacode.com/article/react-typescript-setinterval/
// Start the interval
const startInstanceInterval = (seconds:number) => {
  if (instanceIntervalRef.current !== null) stopInstanceInterval();
  instanceIntervalRef.current = window.setInterval(() => {
    fetchInstance(selectedInstance.instanceId)
  }, seconds*1000);
};

// Stop the interval
const stopInstanceInterval = () => {
  if (instanceIntervalRef.current) {
    window.clearInterval(instanceIntervalRef.current);
    instanceIntervalRef.current = null;
  }
};

// Use the useEffect hook to cleanup the interval when the component unmounts
useEffect(() => {
  // here's the cleanup function
  return () => {
    if (instanceIntervalRef.current !== null) {
      window.clearInterval(instanceIntervalRef.current);
    }
  };
}, []);
  const renderInstanceId = (item?: any, index?: number, column?: IColumn) => {
    return <Link onClick={(ev: React.MouseEvent<unknown>) => {
      fetchInstance(item.instanceId);
    }}    >

      {item.instanceId}</Link>;

  };
  const renderDate = (date: Date) => {

    return format(utcToZonedTime(date, Intl.DateTimeFormat().resolvedOptions().timeZone), 'yyyy-MM-dd HH:mm:ss');

  };
  const renderDateColumn = (item?: any, index?: number, column?: IColumn) => {
    if (item[column.fieldName]) {
      return (renderDate(item[column.fieldName]));
      // return format(utcToZonedTime(item[column.fieldName], Intl.DateTimeFormat().resolvedOptions().timeZone), 'yyyy-MM-dd HH:mm:ss(XX)');
    }
  };
  const renderHistoryName = (item?: IHistoryEvent, index?: number, column?: IColumn) => {

    if (item.EventType === "TaskScheduled") {
      return item.Name;
    }
    else {
      return item.FunctionName;
    }
    // return format(utcToZonedTime(item[column.fieldName], Intl.DateTimeFormat().resolvedOptions().timeZone), 'yyyy-MM-dd HH:mm:ss(XX)');
  }

  const zeroPad = (num: number) => {
    //padstart in rs2016
    const temp = num.toString();
    if (temp.length == 2) { return temp; } else {
      return "0" + temp;
    }


  }
  const renderInstanceDuration = (item?: any, index?: number, column?: IColumn) => {


    if (item.createdTime && item.lastUpdatedTime) {
      const duration = intervalToDuration({
        start: new Date(item.createdTime),
        end: new Date(item.lastUpdatedTime)
      });
      const formatted = [
        duration.hours,
        duration.minutes,
        duration.seconds,
      ]
        .map(zeroPad)
        .join(':')
      return formatted;
      //return formatDuration(intervalToDuration({ start: new Date(item.createdTime), end: new Date(item.lastUpdatedTime) }))
    }
    // return format(utcToZonedTime(item[column.fieldName], Intl.DateTimeFormat().resolvedOptions().timeZone), 'yyyy-MM-dd HH:mm:ss(XX)');

  };
  const renderActivityDuration = (item?: any, index?: number, column?: IColumn) => {
    if (item.ScheduledTime && item.Timestamp) {
      const duration = intervalToDuration({
        start: new Date(item.ScheduledTime),
        end: new Date(item.Timestamp)
      });
      const formatted = [
        duration.hours,
        duration.minutes,
        duration.seconds,
      ]
        .map(zeroPad)
        .join(':')
      return formatted;
      //return formatDuration(intervalToDuration({ start: new Date(item.createdTime), end: new Date(item.lastUpdatedTime) }))
    }
    // return format(utcToZonedTime(item[column.fieldName], Intl.DateTimeFormat().resolvedOptions().timeZone), 'yyyy-MM-dd HH:mm:ss(XX)');

  };
  const instancesCols: IColumn[] = [
    {
      name: 'Instance Id',
      minWidth: 230,
      key: 'instanceId',
      fieldName: 'instanceId',
      isResizable: true,
      onRender: renderInstanceId
    },
    {
      name: 'Name',
      minWidth: 200,
      key: 'name',
      fieldName: 'name',
      isResizable: true,
    }, {
      name: 'Created',
      minWidth: 110,
      key: 'createdTime',
      fieldName: 'createdTime',
      onRender: renderDateColumn,
      isResizable: true,
    },
    {
      name: 'Last Updated',
      minWidth: 110,
      key: 'lastUpdatedTime',
      fieldName: 'lastUpdatedTime',
      onRender: renderDateColumn,
      isResizable: true,
    },
    {
      name: 'Duration',
      minWidth: 60,
      key: 'Duration',
      fieldName: 'Duration',
      onRender: renderInstanceDuration,
      isResizable: true,
    },
    {
      name: 'Custom Status',
      minWidth: 200,
      key: 'customStatus',
      fieldName: 'customStatus',
      isResizable: true,
    },
    {
      name: 'Runtime Status',
      minWidth: 200,
      key: 'runtimeStatus',
      fieldName: 'runtimeStatus',
      isResizable: true,
    },
    {
      name: 'Input',
      minWidth: 200,
      key: 'input',
      fieldName: 'input',
      isResizable: true,
    },
    {
      name: 'Output',
      minWidth: 200,
      key: 'output',
      fieldName: 'output',
      isResizable: true,
    },

  ]
  const historyCols: IColumn[] = [
    {
      name: 'EventType',
      minWidth: 100,
      key: 'EventType',
      fieldName: 'EventType',
      isResizable: true,

    },
    {
      name: 'Name',
      minWidth: 200,
      key: 'Name',
      fieldName: 'Name',
      isResizable: true, onRender: renderHistoryName
    },
    {
      name: 'ScheduledTime',
      minWidth: 110,
      key: 'ScheduledTime',
      fieldName: 'ScheduledTime',
      onRender: renderDateColumn,
      isResizable: true,
    },
    {
      name: 'Timestamp',
      minWidth: 110,
      key: 'Timestamp',
      fieldName: 'Timestamp',
      onRender: renderDateColumn,
      isResizable: true,
    },

    {
      name: 'Duration',
      minWidth: 60,
      key: 'Duration',
      fieldName: 'Duration',
      onRender: renderActivityDuration,
      isResizable: true,
    },
    {
      name: 'Result',
      minWidth: 200,
      key: 'Result',
      fieldName: 'Result',
      isResizable: true,
    },
    {
      name: 'ScheduledTime',
      minWidth: 110,
      key: 'ScheduledTime',
      fieldName: 'ScheduledTime',
      onRender: renderDateColumn,
      isResizable: true,
    },


  ]
  const historyCmds: ICommandBarItemProps[] = [

    {
      name: `Auto Refresh ${refreshDescription?refreshDescription:''}`,
      minWidth: 200,
      key: 'Name',
      fieldName: 'Name',
      checked:refreshSeconds !==null,
      canCheck:true,

      isResizable: true,
      subMenuProps: {
        items: [
          {
            name: "Never", key: "never",
            onClick: () => {
              setRefreshSeconds(null);
            
              setRefreshDescription(null);
            }
          },
          {
            name: "Every 5 Seconds", key: "Every 5 Seconds",
            checked:refreshSeconds===5,
            canCheck:true,
            onClick: () => {
              setRefreshSeconds(5);
              startInstanceInterval(5);
              setRefreshDescription("(Every 5 Seconds)")
            }
          },
          { name: "Every 30 Seconds", key: "Every 30 Seconds",
          checked:refreshSeconds===30,
          canCheck:true,

          onClick: () => {
            setRefreshSeconds(30);
            startInstanceInterval(30);
            setRefreshDescription("(Every 30 Seconds)")
          } },
          { name: "Every Minute", key: "Every Minute",
          checked:refreshSeconds===60,
          canCheck:true,

          onClick: () => {
            setRefreshSeconds(60);
            startInstanceInterval(60);
            setRefreshDescription("(Every Minute)")
          } }

        ]
      }
    },
    {
      name: 'Refresh',
      minWidth: 110,
      key: 'ScheduledTime',
      fieldName: 'ScheduledTime',
      iconProps: { iconName: 'Refresh' },
      isResizable: true,
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        fetchInstance(selectedInstance.instanceId);
      }
    },



  ]
  const historyCmdsFar: ICommandBarItemProps[] = [
    {
      name: 'Back',
      minWidth: 100,
      key: 'EventType',
      fieldName: 'EventType',
      isResizable: true,
      iconProps: { iconName: 'Back' },
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
           stopInstanceInterval();
           setSelectedInstance(null);
      }
    },
 


  ]
  const {
    baseUrl,
    taskHub,
    systemKey, httpClient

  } = props;



  useEffect(() => {

    const fetchData = async () => {
      const url = `${baseUrl}/runtime/webhooks/durableTask/instances?taskHub=${taskHub}&code=${systemKey}`;
      props.httpClient.fetch(url, HttpClient.configurations.v1, {
        headers: { "Accept": "application/json" }
      })
        .then(resp => {

          resp.json().then(instances => {
            setInstances(orderBy(instances, 'createdTime', 'desc'));
          }).catch(e => {
            debugger;
          })

        })
        .catch(e => {
          debugger;
        })
    }

    fetchData();

  }, [])

  
  return (
    <section>

      {selectedInstance &&
        <div>
          <CommandBar items={historyCmds} farItems={historyCmdsFar} />
        
          <div className={styles.grid}>

            <TextField label='Instance Id' value={selectedInstance.instanceId}></TextField>
            <TextField label='Name' value={selectedInstance.name}></TextField>
            <TextField label='Created Time' value={renderDate(selectedInstance.createdTime)}></TextField>
            <TextField label='Last Updated Time' value={renderDate(selectedInstance.lastUpdatedTime)}></TextField>
            <TextField label='Runtime Status' value={selectedInstance.runtimeStatus}></TextField>
            <TextField label='Custom Status' value={selectedInstance.customStatus}></TextField>
            <TextField className={styles.gridFullWidth} label='Output' value={selectedInstance.output}
              multiline={true}
            ></TextField>

          </div>
          <DetailsList items={selectedInstance.historyEvents} columns={historyCols} />
        </div>

      }

      {!selectedInstance &&
        <DetailsList
          items={instances}
          columns={instancesCols}

        />
      }
    </section>
  );

  function fetchInstance(instanceId: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${instanceId}?taskHub=${taskHub}&code=${systemKey}&showHistory=true&showHistoryOutput=true&showInput=true`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      headers: { "Accept": "application/json" }
    })
      .then(resp => {


        resp.json().then(instance => {
      
          setSelectedInstance(instance);
        }).catch(e => {
          debugger;
        });

      })
      .catch(e => {
        debugger;
      });
  }
}

