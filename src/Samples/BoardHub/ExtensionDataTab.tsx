import React, {
  MutableRefObject,
  useEffect,
  useReducer,
  useRef,
  useState,
} from "react";
import * as SDK from "azure-devops-extension-sdk";
import {
  CommonServiceIds,
  IDocumentOptions,
  IExtensionDataManager,
  IExtensionDataService,
} from "azure-devops-extension-api";

import { Button } from "azure-devops-ui/Button";
import { TextField } from "azure-devops-ui/TextField";

interface SimpleDataState {
  type: "simple";
  id: string;
  options?: IDocumentOptions;
  currentText?: string;
  persistedText?: string;
  ready?: boolean;
}

type DataState = SimpleDataState;

type SetSimpleDataState = React.Dispatch<React.SetStateAction<SimpleDataState>>;
type SetDataState = SetSimpleDataState;

function getInitialSimpleDataState(
  id: string,
  options?: IDocumentOptions
): SimpleDataState {
  return Object.freeze({
    type: "simple",
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
    dataStates.map(([state, _]) => {
      return dataManager.getValue<string>(state.id, state.options);
    })
  );

  dataStates.forEach(([_, setState], i) => {
    setState((prevState: any) => ({
      ...prevState,
      ready: true,
      currentText: newValues[i] ?? "",
      persistedText: newValues[i] ?? "",
    }));
  });
}

async function onPersistState(state: DataState, setValue: SetDataState) {
  setValue((prevState: any) => ({ ...prevState, ready: false }));

  const dataManager = await getDataManager();

  let result;
  result = await dataManager.setValue(
    state.id,
    state.currentText || "",
    state.options
  );

  console.log(state.id, "result", result);
  setValue((prevState: any) => ({
    ...prevState,
    ready: true,
    persistedText: state.currentText,
  }));
}

export const ExtensionDataTab: React.FC = () => {
  const [error, setError] = useState<Error | null>(null);
  const [simpleSharedValue, setSimpleSharedValue] = useState(
    getInitialSimpleDataState("simple-shared-value")
  );
  const [simpleUserValue, setSimpleUserValue] = useState(
    getInitialSimpleDataState("simple-user-value", {
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
    <div className="page-content page-content-top rhythm-horizontal-16">
      {error && (
        <>
          <h2>ERROR</h2>
          <div>
            <pre>{error.toString()}</pre>
          </div>
        </>
      )}
      <div className=" flex-row ">
        <EditSimpleDataState
          label={simpleSharedValue.id}
          state={simpleSharedValue}
          setState={setSimpleSharedValue}
          onPersistState={onPersistState}
        />
        <EditSimpleDataState
          label={simpleUserValue.id}
          state={simpleUserValue}
          setState={setSimpleUserValue}
          onPersistState={onPersistState}
        />
      </div>
      <ManageDocumentCollection collectionName="test" />
    </div>
  );
};

const EditSimpleDataState: React.FC<{
  label: string;
  state: SimpleDataState;
  setState: SetSimpleDataState;
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

interface Document {
  id: string;
}

interface ManageCollectionState {
  state: "loading" | "idle" | "edit" | "editing";
  documents: Document[];
  currentDocument: { current: string; persisted?: Document } | null;
  error: Error | null;
}

type LoadComplete = { type: "LOAD_COMPLETE"; documents: Document[] };
type CreateDocument = { type: "CREATE_DOCUMENT" };
type ShowDocument = { type: "SHOW_DOCUMENT"; id: string };
type UpdateCurrent = { type: "UPDATE_CURRENT"; newValue: string };
type SetDocumentStart = { type: "SET_DOCUMENT_START" };
type SetDocumentFinish = { type: "SET_DOCUMENT_FINISH"; newDoc: Document };
type EditError = { type: "EDIT_ERROR"; error: Error };
type Action =
  | LoadComplete
  | CreateDocument
  | ShowDocument
  | UpdateCurrent
  | SetDocumentStart
  | SetDocumentFinish
  | EditError;

const initialState: ManageCollectionState = {
  state: "loading",
  documents: [],
  currentDocument: null,
  error: null,
};

function usePrevious<T>(value: T): T | null {
  const ref = useRef<T>();
  useEffect(() => {
    ref.current = value;
  });
  return ref.current ?? null;
}

function reducer(
  state: ManageCollectionState,
  action: Action
): ManageCollectionState {
  switch (action.type) {
    case "LOAD_COMPLETE":
      return {
        ...state,
        state: "idle",
        documents: action.documents,
        error: null,
      };
    case "CREATE_DOCUMENT":
      return {
        ...state,
        state: "edit",
        currentDocument: {
          current: "",
        },
      };
    case "SHOW_DOCUMENT":
      let newDoc = state.documents.find((doc) => doc.id == action.id);
      if (!newDoc) {
        throw new Error("Could not find document: " + action.id);
      }

      return {
        ...state,
        state: "edit",
        currentDocument: {
          current: JSON.stringify(newDoc, null, 2),
          persisted: newDoc,
        },
        error: null,
      };
    case "UPDATE_CURRENT":
      return {
        ...state,
        currentDocument: {
          ...state.currentDocument,
          current: action.newValue,
        },
        error: null,
      };
    case "SET_DOCUMENT_START":
      return {
        ...state,
        state: "editing",
        error: null,
      };
    case "SET_DOCUMENT_FINISH": {
      let newDoc = action.newDoc;
      let existingDocIndex = state.documents.findIndex(
        (doc) => doc.id === newDoc.id
      );

      let documents;
      if (existingDocIndex == -1) {
        documents = [...state.documents, newDoc];
      } else {
        documents = [
          ...state.documents.slice(0, existingDocIndex),
          newDoc,
          ...state.documents.slice(existingDocIndex + 1),
        ];
      }
      return {
        ...state,
        state: "edit",
        documents,
        currentDocument: {
          current: JSON.stringify(newDoc, null, 2),
          persisted: newDoc,
        },
        error: null,
      };
    }
    case "EDIT_ERROR":
      return {
        ...state,
        error: action.error,
      };
    default:
      return state;
  }
}

const ManageDocumentCollection: React.FC<{
  collectionName: string;
  options?: IDocumentOptions;
}> = ({ collectionName, options }) => {
  const [state, dispatch] = useReducer(reducer, initialState);

  useEffect(() => {
    if (state.state == "loading") {
      (async () => {
        const dataManager = await getDataManager();
        const documents = await dataManager.getDocuments(
          collectionName,
          options
        );
        dispatch({ type: "LOAD_COMPLETE", documents });
      })();
    } else if (state.state == "editing") {
      (async () => {
        const dataManager = await getDataManager();
        let newDoc = await dataManager.setDocument(
          collectionName,
          state.currentDocument?.current,
          options
        );
        dispatch({ type: "SET_DOCUMENT_FINISH", newDoc });
      })();
    }
  }, [state.state]);

  if (state.state == "loading") {
    return (
      <div>Loading documents from the "{collectionName}" collection...</div>
    );
  }

  return (
    <div>
      <h2>{collectionName} Documents</h2>
      <div>
        <label>Select a document</label>
        <select
          value={
            state.currentDocument == null ||
            state.currentDocument.persisted == null
              ? "$create"
              : state.currentDocument.persisted.id
          }
          onChange={(e) => {
            const docId = e.currentTarget.value;
            if (e.currentTarget.value == "$create") {
              dispatch({ type: "CREATE_DOCUMENT" });
            } else {
              dispatch({ type: "SHOW_DOCUMENT", id: docId });
            }
          }}
        >
          {state.documents.map((doc) => (
            <option key={doc.id} value={doc.id}>
              {doc.id}
            </option>
          ))}
          <option key="$create" value="$create">
            Create new document...
          </option>
        </select>
      </div>
      {(state.state == "edit" || state.state == "editing") && (
        <div>
          <textarea
            disabled={state.state == "editing"}
            value={
              state.currentDocument
                ? JSON.stringify(state.currentDocument, null, 2)
                : ""
            }
            onChange={(e) =>
              dispatch({
                type: "UPDATE_CURRENT",
                newValue: e.currentTarget.value,
              })
            }
          ></textarea>
          <Button
            disabled={state.state == "editing"}
            onClick={() => {
              if (!state.currentDocument) {
                dispatch({
                  type: "EDIT_ERROR",
                  error: new Error(
                    "No `currentDocument` set. It is likely a bug!"
                  ),
                });
                return;
              }

              try {
                let newDoc = JSON.parse(state.currentDocument.current);
                if (newDoc.id) {
                  dispatch({ type: "SET_DOCUMENT_START" });
                } else {
                  dispatch({
                    type: "EDIT_ERROR",
                    error: new Error("Document must have an ID property"),
                  });
                }
              } catch (e) {
                dispatch({ type: "EDIT_ERROR", error: e as Error });
              }
            }}
          >
            {state.currentDocument?.persisted == null ? "Create" : "Edit"}
          </Button>
          {state.error && <p>{state.error.toString()}</p>}
        </div>
      )}
    </div>
  );
};
