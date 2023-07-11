import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { IHistoryEvent, IInstance } from '../../../model'
import styles from './DurableFunctionsStatus.module.scss';
import { IDurableFunctionsStatusProps } from './IDurableFunctionsStatusProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http'
import { forEach, groupBy, orderBy, result, sortBy } from 'lodash';
import { Button, DetailsList, Fabric, IColumn, IDetailsRowProps, Link, PrimaryButton, TextField, DetailsRow, CommandBar, ICommandBarItemProps, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { render } from 'react-dom';

import { format, formatDuration, intervalToDuration } from 'date-fns';
import { utcToZonedTime } from 'date-fns-tz';


export default function DurableFunctionsStatus(props: IDurableFunctionsStatusProps): JSX.Element {
  //#region "Hooks"
  const [selectedInstance, setSelectedInstance] = React.useState<IInstance>(null);
  const [instances, setInstances] = useState<Array<IInstance>>([])
  const [refreshSelectedInstanceSeconds, setRefreshSelectedInstanceSeconds] = useState<number>(null);
  const [refreshSelectedInstanceDescription, setRefreshSelectedInstanceDescription] = useState<string>(null);
  const selectedInstanceIntervalRef = useRef<number | null>(null);
  const [refreshInstancesSeconds, setRefreshInstancesSeconds] = useState<number>(null);
  const [refreshInstancesDescription, setRefreshInstancesDescription] = useState<string>(null);
  const instancesIntervalRef = useRef<number | null>(null);
  const [alertMessage, setAlertMessage] = useState<{ message: string, type: MessageBarType } | null>(null);
  useEffect(() => {
    fetchInstancesData();

  }, []);

  //#endregion "Hooks"

  //#region Timer functions
  //* See https://www.kindacode.com/article/react-typescript-setinterval/
  // Start the interval
  const startSelectedInstanceInterval = (seconds: number) => {
    if (selectedInstanceIntervalRef.current !== null) stopSelectedInstanceInterval();
    selectedInstanceIntervalRef.current = window.setInterval(() => {
      fetchSelectedInstance(selectedInstance.instanceId)
    }, seconds * 1000);
  };

  // Stop the interval
  const stopSelectedInstanceInterval = () => {
    if (selectedInstanceIntervalRef.current) {
      window.clearInterval(selectedInstanceIntervalRef.current);
      selectedInstanceIntervalRef.current = null;
    }
  };
  const startInstancesInterval = (seconds: number) => {
    if (instancesIntervalRef.current !== null) stopInstancesInterval();
    instancesIntervalRef.current = window.setInterval(() => {
      fetchInstancesData()
    }, seconds * 1000);
  };

  // Stop the interval
  const stopInstancesInterval = () => {
    if (instancesIntervalRef.current) {
      window.clearInterval(instancesIntervalRef.current);
      instancesIntervalRef.current = null;
    }
  };


  // Use the useEffect hook to cleanup the interval when the component unmounts
  useEffect(() => {
    // here's the cleanup function
    return () => {
      if (selectedInstanceIntervalRef.current !== null) {
        window.clearInterval(selectedInstanceIntervalRef.current);
        window.clearInterval(instancesIntervalRef.current);
      }
    };
  }, []);

  //#endregion Timer functions
  //#region Render methods

  const getUniqueStatuses = (): Array<string> => {
    debugger;
    var stati = []
    const groups = groupBy(instances, 'runtimeStatus')
    for (const group in groups) {
      stati.push(group)
    }
    return stati;
  }


  const renderInstanceId = (item?: any, index?: number, column?: IColumn) => {
    return <Link onClick={(ev: React.MouseEvent<unknown>) => {
      fetchSelectedInstance(item.instanceId);
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
  const renderOutput = (item?: any, index?: number, column?: IColumn) => {


    debugger;
    if (item[column.fieldName]) {
      return item[column.fieldName].toString();
    } 
    else {
      return null;
    }

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
  //#endregion Render methods
  //#region data 
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
      onRender: renderOutput,
    },

  ]
  const instanceCmds: ICommandBarItemProps[] = [

    {
      name: `Auto Refresh ${refreshInstancesDescription ? refreshInstancesDescription : ''}`,
      key: 'Name',
      subMenuProps: {
        items: [
          {
            name: "Never", key: "never",
            checked: refreshInstancesSeconds === null,
            canCheck: true,
            onClick: () => {
              setRefreshInstancesSeconds(null);
              setRefreshInstancesDescription(null);
            }
          },
          {
            name: "Every 5 Seconds", key: "Every 5 Seconds",
            checked: refreshInstancesSeconds === 5,
            canCheck: true,
            onClick: () => {
              setRefreshInstancesSeconds(5);
              startInstancesInterval(5);
              setRefreshInstancesDescription("(Every 5 Seconds)")
            }
          },
          {
            name: "Every 30 Seconds", key: "Every 30 Seconds",
            checked: refreshInstancesSeconds === 30,
            canCheck: true,

            onClick: () => {
              setRefreshInstancesSeconds(30);
              startInstancesInterval(30);
              setRefreshInstancesDescription("(Every 30 Seconds)")
            }
          },
          {
            name: "Every Minute", key: "Every Minute",
            checked: refreshInstancesSeconds === 60,
            canCheck: true,
            onClick: () => {
              setRefreshInstancesSeconds(60);
              startInstancesInterval(60);
              setRefreshInstancesDescription("(Every Minute)")
            }
          }
        ]
      }
    },
    {
      name: 'Refresh',
      key: 'ScheduledTime',
      iconProps: { iconName: 'Refresh' },
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        fetchInstancesData();
      }
    },
    {
      name: 'Purge History',
      key: 'Purge',
      iconProps: { iconName: 'Delete' },
      subMenuProps: {
        items: getUniqueStatuses().map(status => {
          return {
            name: status,
            key: status,
            onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
              purgeByStatus(status);
            }
          }
        })
      }
    },
    {
      name: 'Start New Orchestration',
      key: 'New',
      iconProps: { iconName: 'Add' },

      subMenuProps: {
        items: props.orchestrationNames.map((orchName => {
          return {
            name: orchName,
            key: orchName,
            onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
              startOrchestration(orchName);
              fetchInstancesData();
            }
          }
        }))
      }
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
      name: `Auto Refresh ${refreshSelectedInstanceDescription ? refreshSelectedInstanceDescription : ''}`,
      key: 'Name',
      checked: refreshSelectedInstanceSeconds !== null,
      canCheck: true,
      subMenuProps: {
        items: [
          {
            name: "Never", key: "never",
            onClick: () => {
              setRefreshSelectedInstanceSeconds(null);
              setRefreshSelectedInstanceDescription(null);
            }
          },
          {
            name: "Every 5 Seconds", key: "Every 5 Seconds",
            checked: refreshSelectedInstanceSeconds === 5,
            canCheck: true,
            onClick: () => {
              setRefreshSelectedInstanceSeconds(5);
              startSelectedInstanceInterval(5);
              setRefreshSelectedInstanceDescription("(Every 5 Seconds)")
            }
          },
          {
            name: "Every 30 Seconds", key: "Every 30 Seconds",
            checked: refreshSelectedInstanceSeconds === 30,
            canCheck: true,

            onClick: () => {
              setRefreshSelectedInstanceSeconds(30);
              startSelectedInstanceInterval(30);
              setRefreshSelectedInstanceDescription("(Every 30 Seconds)")
            }
          },
          {
            name: "Every Minute", key: "Every Minute",
            checked: refreshSelectedInstanceSeconds === 60,
            canCheck: true,
            onClick: () => {
              setRefreshSelectedInstanceSeconds(60);
              startSelectedInstanceInterval(60);
              setRefreshSelectedInstanceDescription("(Every Minute)")
            }
          }
        ]
      }
    },
    {
      name: 'Refresh',
      key: 'ScheduledTime',
      iconProps: { iconName: 'Refresh' },
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        fetchSelectedInstance(selectedInstance.instanceId);
      }
    },
    {
      name: 'Purge',
      key: 'Purge',
      iconProps: { iconName: 'Delete' },
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        purgeSelectedInstance(selectedInstance.instanceId)
      }
    },
    {
      name: 'Terminate',
      key: 'Terminate',
      disabled: (selectedInstance && (selectedInstance.runtimeStatus === "Terminated" || selectedInstance.runtimeStatus === "Completed")),
      iconProps: { iconName: 'Stop' },
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        debugger;
        terminateSelectedInstance(selectedInstance.instanceId);
        fetchInstancesData()
      }
    },
    {
      name: 'Suspend',
      key: 'Suspend',
      iconProps: { iconName: 'Pause' },
      disabled: (selectedInstance && (selectedInstance.runtimeStatus === "Terminated" || selectedInstance.runtimeStatus === "Completed" || selectedInstance.runtimeStatus === "Suspended")),

      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        debugger;
        suspendSelectedInstance(selectedInstance.instanceId)
      }
    },
    {
      name: 'Resume',
      key: 'Resume',
      iconProps: { iconName: 'Play' },
      disabled: (selectedInstance && (selectedInstance.runtimeStatus !== "Suspended")),

      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        debugger;
        resumeSelectedInstance(selectedInstance.instanceId)
      }
    },
  ]
  const historyCmdsFar: ICommandBarItemProps[] = [
    {
      name: 'Back',
      key: 'EventType',
      iconProps: { iconName: 'Back' },
      onClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined) => {
        stopSelectedInstanceInterval();
        setSelectedInstance(null);
      }
    },
  ]

  const {
    baseUrl,
    taskHub,
    systemKey, httpClient

  } = props;
  //#endregion data 
  //#region IO
  const fetchInstancesData = async () => {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances?taskHub=${taskHub}&code=${systemKey}`;
    props.httpClient.fetch(url, HttpClient.configurations.v1, {
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        return resp.json()
      })
      .then(instances => {
        setInstances(orderBy(instances, 'createdTime', 'desc'));
      })
      .catch(e => {
        debugger;
      })
  }
  function purgeSelectedInstance(instanceId: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${instanceId}?taskHub=${taskHub}&code=${systemKey}`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      method: "DELETE",
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        setSelectedInstance(null);
        fetchInstancesData();
      })
      .catch(e => {
        debugger;
      });
  }
  function purgeByStatus(status: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances?taskHub=${taskHub}&code=${systemKey}&createdTimeFrom=2023-07-10&runtimeStatus=${status}`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      method: "DELETE",
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        return resp.json()

      }).then(e => {
        debugger;
        alert(`${e.instancesDeleted} instances deleted.`)
        setTimeout(() => {
          fetchInstancesData();
        }, 500)

      })
      .catch(e => {
        debugger;
      });
  }
  function terminateSelectedInstance(instanceId: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${instanceId}/terminate?taskHub=${taskHub}&code=${systemKey}`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      method: "Post",
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        switch (resp.status) {
          case 404:
            alert("This instance was not found");
            break;
          case 410:
            alert("This instance has completed or failed");
            break;
        }
      })
      .then(() => {
        setTimeout(() => {
          fetchSelectedInstance(selectedInstance.instanceId);
        }, 500)

      })
      .catch(e => {
        debugger;
      });
  }
  function suspendSelectedInstance(instanceId: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${instanceId}/suspend?taskHub=${taskHub}&code=${systemKey}`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      method: "Post",
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        switch (resp.status) {
          case 404:
            alert("This instance was not found");
            break;
          case 410:
            alert("This instance has completed or failed");
            break;
        }
      })
      .then(() => {
        setTimeout(() => {
          fetchSelectedInstance(selectedInstance.instanceId);
        }, 500)

      })
      .catch(e => {
        debugger;
      });
  }
  function resumeSelectedInstance(instanceId: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${instanceId}/resume?taskHub=${taskHub}&code=${systemKey}`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      method: "Post",
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        switch (resp.status) {
          case 404:
            alert("This instance was not found");
            break;
          case 410:
            alert("This instance has completed or failed");
            break;
        }
      })
      .then(() => {
        setTimeout(() => {
          fetchSelectedInstance(selectedInstance.instanceId);
        }, 500)

      })
      .catch(e => {
        debugger;
      });
  }
  function startOrchestration(orchestrationName: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/orchestrators/${orchestrationName}?taskHub=${taskHub}&code=${systemKey}`;
    debugger;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      method: "POST",
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        return resp.json()

      })
      .then((msg) => {
        debugger;
        setAlertMessage({ message: `Instance ${msg.id} started.`, type: MessageBarType.success });
        fetchInstancesData();

      })
      .catch(e => {
        debugger;
      });
  }
  function fetchSelectedInstance(instanceId: string) {
    const url = `${baseUrl}/runtime/webhooks/durableTask/instances/${instanceId}?taskHub=${taskHub}&code=${systemKey}&showHistory=true&showHistoryOutput=true&showInput=true`;
    httpClient.fetch(url, HttpClient.configurations.v1, {
      headers: { "Accept": "application/json" }
    })
      .then((resp: HttpClientResponse) => {
        return resp.json()
      })
      .then(instance => {
        setSelectedInstance(instance);
      })
      .catch(e => {
        debugger;
      });
  }
  //#endregion IO
  //const messages:Array<MessageBar>=[]
  return (
    <section>
      {
        alertMessage &&
        <div>
          <MessageBar messageBarType={alertMessage.type} onDismiss={() => { setAlertMessage(null) }}> {alertMessage.message}</MessageBar>
        </div>
      }

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
        <div>
          <CommandBar items={instanceCmds} />

          <DetailsList
            items={instances}
            columns={instancesCols}

          />
        </div>
      }
    </section>
  );
}


