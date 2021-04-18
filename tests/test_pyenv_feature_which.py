from tempenv import TemporaryEnvironment
from test_pyenv import TestPyenvBase
from test_pyenv_helpers import run_pyenv_test


class TestPyenvFeatureWhich(TestPyenvBase):
    def test_which_exists_is_global(self, setup):
        def commands(ctx):
            for name in ['python', 'python3', 'python38', 'pip3', 'pip3.8']:
                sub_dir = '' if 'python' in name else 'Scripts\\'
                assert ctx.pyenv(["which", name]) == rf'{ctx.pyenv_path}\versions\3.8.5\{sub_dir}{name}.exe' or \
                    ctx.pyenv(["which", name]) == rf'{ctx.pyenv_path}\versions\3.8.5\{sub_dir}{name}.EXE'
        settings = {
            'versions': ['3.8.5'],
            'global_ver': '3.8.5'
        }
        run_pyenv_test(settings, commands)

    def test_which_exists_is_local(self, setup):
        def commands(ctx):
            for name in ['python', 'python3', 'python38', 'pip3', 'pip3.8']:
                sub_dir = '' if 'python' in name else 'Scripts\\'
                assert ctx.pyenv(["which", name]) == rf'{ctx.pyenv_path}\versions\3.8.5\{sub_dir}{name}.exe' or \
                    ctx.pyenv(["which", name]) == rf'{ctx.pyenv_path}\versions\3.8.5\{sub_dir}{name}.EXE'
        settings = {
            'versions': ['3.8.5'],
            'local_ver': '3.8.5'
        }
        run_pyenv_test(settings, commands)

    def test_which_exists_is_shell(self, setup):
        def commands(ctx):
            for name in ['python', 'python3', 'python38', 'pip3', 'pip3.8']:
                sub_dir = '' if 'python' in name else 'Scripts\\'
                assert ctx.pyenv(["which", name]) == rf'{ctx.pyenv_path}\versions\3.8.5\{sub_dir}{name}.exe' or \
                    ctx.pyenv(["which", name]) == rf'{ctx.pyenv_path}\versions\3.8.5\{sub_dir}{name}.EXE'
        with TemporaryEnvironment({"PYENV_VERSION": "3.8.5"}):
            run_pyenv_test({'versions': ['3.8.5']}, commands)

    def test_which_exists_is_global_not_installed(self, setup):
        def commands(ctx):
            for name in ['python', 'python3', 'python38', 'pip3', 'pip3.8']:
                assert ctx.pyenv(["which", name]) == "pyenv: version `3.8.5' is not installed (set by 3.8.5)"
        run_pyenv_test({'global_ver': '3.8.5'}, commands)

    def test_which_exists_is_local_not_installed(self, setup):
        def commands(ctx):
            for name in ['python', 'python3', 'python38', 'pip3', 'pip3.8']:
                assert ctx.pyenv(["which", name]) == "pyenv: version `3.8.5' is not installed (set by 3.8.5)"
        run_pyenv_test({'local_ver': '3.8.5'}, commands)

    def test_which_exists_is_shell_not_installed(self, setup):
        def commands(ctx):
            for name in ['python', 'python3', 'python38', 'pip3', 'pip3.8']:
                assert ctx.pyenv(["which", name]) == "pyenv: version `3.8.5' is not installed (set by 3.8.5)"
        with TemporaryEnvironment({"PYENV_VERSION": "3.8.5"}):
            run_pyenv_test({}, commands)

    def test_which_exists_is_global_other_version(self, setup):
        def commands(ctx):
            for name in ['python38', 'pip3.8']:
                assert ctx.pyenv(["which", name]) == (f"pyenv: {name}: command not found\r\n"
                                                      f"\r\n"
                                                      f"The '{name}' command exists in these Python versions:\r\n"
                                                      f"  3.8.2\r\n"
                                                      f"  3.8.6\r\n"
                                                      f"  ")
        settings = {
            'versions': ['3.8.2', '3.8.6', '3.9.1'],
            'global_ver': '3.9.1'
        }
        run_pyenv_test(settings, commands)

    def test_which_exists_is_local_other_version(self, setup):
        def commands(ctx):
            for name in ['python38', 'pip3.8']:
                assert ctx.pyenv(["which", name]) == (f"pyenv: {name}: command not found\r\n"
                                                      f"\r\n"
                                                      f"The '{name}' command exists in these Python versions:\r\n"
                                                      f"  3.8.2\r\n"
                                                      f"  3.8.6\r\n"
                                                      f"  ")
        settings = {
            'versions': ['3.8.2', '3.8.6', '3.9.1'],
            'local_ver': '3.9.1'
        }
        run_pyenv_test(settings, commands)

    def test_which_exists_is_shell_other_version(self, setup):
        def commands(ctx):
            for name in ['python38', 'pip3.8']:
                assert ctx.pyenv(["which", name]) == (f"pyenv: {name}: command not found\r\n"
                                                      f"\r\n"
                                                      f"The '{name}' command exists in these Python versions:\r\n"
                                                      f"  3.8.2\r\n"
                                                      f"  3.8.6\r\n"
                                                      f"  ")
        settings = {
            'versions': ['3.8.2', '3.8.6', '3.9.1'],
        }
        with TemporaryEnvironment({"PYENV_VERSION": "3.9.1"}):
            run_pyenv_test(settings, commands)

    def test_which_command_not_found(self, setup):
        def commands(ctx):
            for name in ['python3.8']:
                assert ctx.pyenv(["which", name]) == f"pyenv: {name}: command not found"
        settings = {
            'versions': ['3.8.6'],
            'global_ver': '3.8.6'
        }
        run_pyenv_test(settings, commands)

    def test_which_no_version_defined(self, setup):
        def commands(ctx):
            for name in ['python']:
                assert ctx.pyenv(["which", name]) == ("No global python version has been set yet. "
                                                      "Please set the global version by typing:\r\n"
                                                      "pyenv global 3.7.2")
        run_pyenv_test({'versions': ['3.8.6']}, commands)
