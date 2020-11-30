# import abc
#
# class data_abc(metaclass=abc.ABCMeta):
#
#     @abc.abstractmethod
#     def random(self):

import random


class tem_random:

    def __init__(self, build):
        self.low = build.low
        self.high = build.high
        self.accuracy = build.accuracy

    def random(self):
        return round(random.uniform(self.low, self.high), self.accuracy)

    class Build:

        def __init__(self):
            self.low = 36.1
            self.high = 36.8
            self.accuracy = 1

        def set_low(self, low):
            self.low = low
            return self

        def set_high(self, high):
            self.high = high
            return self

        def set_accuracy(self, accuracy):
            self.accuracy = accuracy
            return self

        def build(self):
            return tem_random(self)
