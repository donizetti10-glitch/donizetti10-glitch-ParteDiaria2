var e = {
    button: `Button`,
    checkBox: `Checkbox`,
    comboBox: `Drop`,
    listBox: `List`,
    optionButton: `Radio`,
    groupBox: `GBox`,
    label: `Label`,
    scrollBar: `Scroll`,
    spinner: `Spin`,
    dialog: `Dialog`,
  },
  t = { ...e, checkBox: `CheckBox`, dialog: `Dialog` },
  n = { ...Object.fromEntries(Object.entries(e).map(([e, t]) => [t, e])), CheckBox: `checkBox` },
  r = { unchecked: `0`, checked: `1`, mixed: `2` },
  i = { 0: `unchecked`, 1: `checked`, 2: `mixed` };
export { e as a, n as i, i as n, t as r, r as t };
