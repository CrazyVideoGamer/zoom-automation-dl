This is for myself, it is not related to this repository at all :)

pywingui reference:
* use Accessibility Insights For Windows to locate elements on windows gui (other options: https://pywinauto.readthedocs.io/en/latest/getting_started.html#gui-objects-inspection-spy-tools)
  * **IMPORTANT**: the class name in the inspector may not be the same in pywingui terminlogy; use print_control_identifiers to check.
* use child_elements (more options) or getattribute (dialog["whatever"] -> .child_elements(title="whatever")) to locate elements
  * options for child_elements, when multiple elements have that text:
  * Note: "control type" in windows terminlogy is the same as "class" in pywingui terminlogy
* class_name Elements with this window class
* class_name_re Elements whose class matches this regular expression
* parent Elements that are children of this
* process Elements running in this process
* title Elements with this text
* title_re Elements whose text matches this regular expression
* top_level_only Top level elements only (default=**True**)
* visible_only Visible elements only (default=**True**)
* enabled_only Enabled elements only (default=False)
* best_match Elements with a title similar to this
* handle The handle of the element to return
* ctrl_index The index of the child element to return
* found_index The index of the filtered out child element to return
* predicate_func A user provided hook for a custom element validation
* active_only Active elements only (default=False)
* control_id Elements with this control id
* control_type Elements with this control type (string; for UIAutomation elements)
* auto_id Elements with this automation id (for UIAutomation elements)
* framework_id Elements with this framework id (for UIAutomation elements)
* backend Back-end name to use while searching (default=None means current active backend)
(https://pywinauto.readthedocs.io/en/latest/code/pywinauto.findwindows.html#pywinauto.findwindows.find_elements)
- Control type docs: https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltypesoverview#current-ui-automation-control-types
- Methods for clicking, typing, etc.: https://pywinauto.readthedocs.io/en/latest/controls_overview.html#all-controls