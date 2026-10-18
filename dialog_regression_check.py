"""GUI smoke check: run with a display and the application's dependencies."""
import tempfile
import time
import traceback
from pathlib import Path
from unittest.mock import patch

import main


def pump(app, condition):
    deadline = time.monotonic() + 5
    while not condition():
        if time.monotonic() > deadline:
            raise AssertionError("Dialog did not become ready within five seconds")
        app.update()
        time.sleep(0.01)


with tempfile.TemporaryDirectory() as directory:
    main.PROJECTS_ROOT = Path(directory)
    main.DB_PATH = Path(directory) / "projects.db"
    app = main.CertificateStudioApp()
    errors = []
    app.report_callback_exception = lambda *args: errors.append(''.join(traceback.format_exception(*args)))
    try:
        app.update()
        for scale in (1.0, 1.5, 2.0, 3.0):
            app.tk.call("tk", "scaling", scale)
            dialog = main.ProjectNameDialog(app.launcher)
            pump(app, lambda: app.grab_current() == dialog)
            shell = dialog.winfo_children()[0]
            for widget in shell.winfo_children():
                assert widget.winfo_viewable()
                assert widget.winfo_rooty() + widget.winfo_height() <= dialog.winfo_rooty() + dialog.winfo_height()
                for button in widget.winfo_children():
                    if isinstance(button, main.RoundedButton):
                        assert button.winfo_height() >= button.winfo_reqheight(), (
                            scale, button._text, button.winfo_height(), button.winfo_reqheight())
            dialog.entry.insert(0, "Regression check")
            dialog._submit()
            assert dialog.result == "Regression check"
            assert app.grab_current() is None
            print(f"PASS: layout, modal grab, submission at scale {scale}", flush=True)

        # Reproduce a window manager rejecting grabs after reporting visibility.
        original_grab = main.ProjectNameDialog.grab_set
        attempts = []
        def delayed_grab(dialog):
            attempts.append(True)
            if len(attempts) < 3:
                raise main.tk.TclError("grab failed: window not viewable")
            original_grab(dialog)
        with patch.object(main.ProjectNameDialog, "grab_set", delayed_grab):
            dialog = main.ProjectNameDialog(app.launcher)
            pump(app, lambda: app.grab_current() == dialog)
            assert len(attempts) == 3
            dialog.destroy()
        print("PASS: recovers from two rejected grabs", flush=True)

        # A font/layout change after construction must not squeeze the buttons.
        dialog = main.ProjectNameDialog(app.launcher)
        shell = dialog.winfo_children()[0]
        for widget in shell.winfo_children():
            if isinstance(widget, main.tk.Label):
                widget.configure(font=(main.UI_FONT, 24))
        pump(app, lambda: app.grab_current() == dialog)
        app.update()
        for widget in shell.winfo_children():
            for button in widget.winfo_children():
                if isinstance(button, main.RoundedButton):
                    assert button.winfo_viewable()
                    assert button.winfo_height() >= button.winfo_reqheight(), (
                        "late layout change", button.winfo_height(), button.winfo_reqheight())
        dialog.destroy()
        print("PASS: late font/layout changes preserve button height", flush=True)

        dialog = main.ProjectNameDialog(app.launcher)
        dialog.withdraw()
        app.update()
        assert dialog._modal_after is not None
        dialog.destroy()
        app.update()
        assert app.grab_current() is None
        assert not errors, '\n'.join(errors)
        print("PASS: close before mapping cancels pending activation; no callback errors", flush=True)
    finally:
        app.pm.conn.close()
        app.destroy()
