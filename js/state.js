const defaultState = Object.freeze({
  fileName: "",
  workbook: null,
  sheets: [],
  selectedSheetName: "",
  tableRows: [],
  tableColumns: [],
  visualization: null,
  error: null,
  status: "idle",
});

export function createState(initialState = {}) {
  let state = {
    ...defaultState,
    ...initialState,
  };

  const listeners = new Set();

  const getState = () => state;

  const setState = (updater) => {
    const nextState =
      typeof updater === "function" ? updater(state) : { ...state, ...updater };

    state = nextState;
    listeners.forEach((listener) => listener(state));
    return state;
  };

  const resetState = () => {
    state = {
      ...defaultState,
      ...initialState,
    };

    listeners.forEach((listener) => listener(state));
    return state;
  };

  const subscribe = (listener) => {
    listeners.add(listener);
    listener(state);

    return () => {
      listeners.delete(listener);
    };
  };

  return {
    getState,
    setState,
    resetState,
    subscribe,
  };
}

export { defaultState };
