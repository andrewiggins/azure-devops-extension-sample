import React, { MutableRefObject, useEffect, useRef, useState } from "react";
import * as SDK from "azure-devops-extension-sdk";
import {
  CommonServiceIds,
  IDocumentOptions,
  IExtensionDataManager,
  IExtensionDataService,
} from "azure-devops-extension-api";

import { Button } from "azure-devops-ui/Button";
import { TextField } from "azure-devops-ui/TextField";

export interface DataState {
  id: string;
  options?: IDocumentOptions;
  currentText?: string;
  persistedText?: string;
  ready?: boolean;
}

type SetDataState = React.Dispatch<React.SetStateAction<DataState>>;

function getDefaultDataState(
  id: string,
  options?: IDocumentOptions
): DataState {
  return Object.freeze({
    id,
    options,
    currentText: undefined,
    persistedText: undefined,
    ready: false,
  });
}

const _dataManagerRef: MutableRefObject<IExtensionDataManager | null> = {
  current: null,
};
async function getDataManager() {
  if (!_dataManagerRef.current) {
    await SDK.ready();
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>(
      CommonServiceIds.ExtensionDataService
    );
    _dataManagerRef.current = await extDataService.getExtensionDataManager(
      SDK.getExtensionContext().id,
      accessToken
    );
  }

  return _dataManagerRef.current;
}

async function getInitialData(dataStates: Array<[DataState, SetDataState]>) {
  const dataManager = await getDataManager();
  const newValues = await Promise.all(
    dataStates.map((dataState) => {
      return dataManager.getValue<string>(
        dataState[0].id,
        dataState[0].options
      );
    })
  );

  dataStates.forEach(([state, setState], i) => {
    setState({
      ...state,
      ready: true,
      currentText: newValues[i] ?? "",
      persistedText: newValues[i] ?? "",
    });
  });
}

async function onPersistState(state: DataState, setValue: SetDataState) {
  setValue({ ...state, ready: false });

  const dataManager = await getDataManager();
  await dataManager.setValue(state.id, state.currentText || "", state.options);

  setValue({
    ...state,
    ready: true,
    persistedText: state.currentText,
  });
}

export const ExtensionDataTab: React.FC = () => {
  const [error, setError] = useState<Error | null>(null);
  const [simpleSharedValue, setSimpleSharedValue] = useState(
    getDefaultDataState("simple-shared-value")
  );
  const [simpleUserValue, setSimpleUserValue] = useState(
    getDefaultDataState("simple-user-value", {
      scopeType: "User",
      scopeValue: "Me",
    })
  );

  useEffect(() => {
    getInitialData([
      [simpleSharedValue, setSimpleSharedValue],
      [simpleUserValue, setSimpleUserValue],
    ]).catch((error) => setError(error));
  }, []);

  return (
    <div className="page-content page-content-top flex-row rhythm-horizontal-16">
      {error && (
        <>
          <h2>ERROR</h2>
          <div>
            <pre>{error.toString()}</pre>
          </div>
        </>
      )}
      <EditDataState
        label="Simple shared value"
        state={simpleSharedValue}
        setState={setSimpleSharedValue}
        onPersistState={onPersistState}
      />
      <EditDataState
        label="Simple user value"
        state={simpleUserValue}
        setState={setSimpleUserValue}
        onPersistState={onPersistState}
      />
    </div>
  );
};

const EditDataState: React.FC<{
  label: string;
  state: DataState;
  setState: SetDataState;
  onPersistState: (state: DataState, setState: SetDataState) => void;
}> = ({ label, state, setState, onPersistState }) => {
  return (
    <div>
      <h2>{label}</h2>
      <TextField
        value={state.currentText}
        onChange={(_, newValue) =>
          setState({ ...state, currentText: newValue })
        }
        disabled={!state.ready}
      />
      <Button
        text="Save"
        primary={true}
        onClick={() => onPersistState(state, setState)}
        disabled={!state.ready || state.currentText === state.persistedText}
      />
    </div>
  );
};
