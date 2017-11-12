from subprocess import Popen


class DependencyInstaller(object):

    def __init__(self):
        self._requirements_path = "requirements.txt"
        self._wheels_path = "wheels/"
        self._pip_install_command = "pip install --no-index --find-links=file:{} -r {}".\
            format(self._wheels_path, self._requirements_path)

    def install_dependencies(self):
        command = Popen(self._pip_install_command)
        command.communicate()
        command.kill()

if __name__ == '__main__':
    DependencyInstaller().install_dependencies()
