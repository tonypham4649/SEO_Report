from abc import ABC, abstractmethod


class Command(ABC):

    @abstractmethod
    def execute(self) -> None:
        pass


class Invoker:
    def runCmd(self, command: Command):
        command.execute()


class Receiver():

    def do_st(self) -> None:
        print(f'layer 3 -- run')

    def do_st_else(self) -> None:
        print(f'layer 3 -- run 2')


class getFile(Command):
    def __init__(self, receiver: Receiver):
        self._receiver = receiver

    def execute(self) -> None:
        print(f'layer 2: navigate cmd to layer 3')
        self._receiver.do_st()


class login2SS(Command):
    def __init__(self, receiver: Receiver, loginID: str, loginPass: str) -> None:
        self._loginID = loginID
        self._loginPass = loginPass
        self._receiver = receiver

    def execute(self) -> None:
        print('layer 2: navigate back to receiver class')
        self._receiver.do_st_else()


if __name__ == "__main__":
    invoker = Invoker()
    receiver = Receiver()

    cmd_list = []
    cmd_list.append(getFile(receiver))
    cmd_list.append(login2SS(receiver, 'bao', 'pass'))

    for cmd in cmd_list:
        invoker.runCmd(cmd)
