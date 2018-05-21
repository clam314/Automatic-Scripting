# -*- coding: utf-8 -*-

import configparser


def empty2None(s):
    if s == 'empty' or s == '':
        return None
    else:
        return s


class ConfigHandle(object):
    def __init__(self, config_path):
        cf = configparser.ConfigParser()
        cf.read(config_path, encoding='utf-8')
        self._config = cf

    def getEmailInfo(self):
        section = self._config['smtp']
        return section['host'], section['port'], section['ssl_port']

    def getUser(self):
        section = self._config['user_info']
        return section['user'], section['pwd'], section['nick']

    def getDailyReceiver(self):
        section = self._config['daily_member']
        return empty2None(section['main_receiver']), empty2None(
            section['secondary_receiver'])

    def getExpReceiver(self):
        section = self._config['exp_member']
        return empty2None(section['main_receiver']), empty2None(
            section['secondary_receiver'])

    def getSynExpReceiver(self):
        section = self._config['syn_exp_member']
        return empty2None(section['main_receiver']), empty2None(
            section['secondary_receiver'])
