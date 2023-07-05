import * as React from 'react';
import { useState, useEffect } from 'react';
import { IInstance } from '../../../model'
import styles from './DurableFunctionsStatus.module.scss';
import { IDurableFunctionsStatusProps } from './IDurableFunctionsStatusProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient } from '@microsoft/sp-http'
import { orderBy, result, sortBy } from 'lodash';
import { Button, DetailsList, Fabric, IColumn, IDetailsRowProps, Link, PrimaryButton, TextField, DetailsRow } from 'office-ui-fabric-react';
import { render } from 'react-dom';

import { format, formatDuration, intervalToDuration } from 'date-fns';
import { utcToZonedTime } from 'date-fns-tz';

export default function DurableFunctionsStatus(props: IDurableFunctionsStatusProps): JSX.Element {
  const [selectedInstance, setSelectedInstance] = React.useState<IInstance>(null);
  const [instances, setInstances] = useState<Array<IInstance>>([])

  const renderInstanceId = (item?: any, index?: number, column?: IColumn) => {
    return <Link onClick={(ev: React.MouseEvent<unknown>) => {
      const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${item.instanceId}?taskHub=${taskHub}&code=${systemKey}&showHistory=true&showHistoryOutput=true&showInput=true`;
      httpClient.fetch(url, HttpClient.configurations.v1, {
        headers: { "Accept": "application/json" }
      })
        .then(resp => {


          resp.json().then(x => {

            setSelectedInstance(x);
          }).catch(e => {
            debugger;
          })

        })
        .catch(e => {
          debugger;
        })
    }}>

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
  const zeroPad = (num:number) => {
    //padstart in rs2016
    const temp=num.toString();
    if(temp.length ==2 ) {return temp;}else{
      return "0"+temp;
    }
    
   
}
  const renderInstanceDuration = (item?: any, index?: number, column?: IColumn) => {
    debugger;

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
    debugger;

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
      name: 'FunctionName',
      minWidth: 200,
      key: 'FunctionName',
      fieldName: 'FunctionName',
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

  // useEffect(() => {
  //   debugger;
  //   const fetchData = async () => {
  //     const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${selectedInstanceId}?taskHub=${taskHub}&code=${systemKey}&showHistory=true&showHistoryOutput=true&showInput=true`;
  //     props.httpClient.fetch(url, HttpClient.configurations.v1, {
  //       headers: { "Accept": "application/json" }
  //     })
  //       .then(resp => {


  //         resp.json().then(x => {
  //           debugger;
  //           setSelectedInstance(x);
  //         }).catch(e => {
  //           debugger;
  //         })

  //       })
  //       .catch(e => {
  //         debugger;
  //       })
  //   }

  //   fetchData();

  // }, [selectedInstanceId]);
  debugger;
  return (
    <section>

      {selectedInstance &&
        <div>
          <PrimaryButton onClick={(e) => {
            setSelectedInstance(null);
          }}>Back</PrimaryButton>
          <div className={styles.grid}>

            <TextField label='Instance Id' value={selectedInstance.instanceId}></TextField>
            <TextField label='Name' value={selectedInstance.name}></TextField>
            <TextField label='Created Time' value={renderDate(selectedInstance.createdTime)}></TextField>
            <TextField label='Last Updated Time' value={renderDate(selectedInstance.lastUpdatedTime)}></TextField>
            <TextField label='Runtime Status' value={selectedInstance.runtimeStatus}></TextField>

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
}

