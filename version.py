#!/usr/bin/env python
# -*- coding: utf-8 -*-

from math import copysign

__whats_new = list()


def add_release(ver, whats_new):
    __whats_new.append((ver, whats_new))


def version():
    return __whats_new[-1][0]


def version_cmp(a, b):
    if a == b:
        return 0
    a = float(a or '-1.0')
    b = float(b or '-1.0')
    return int(copysign(1, a-b))


def iterate_whats_new(last_version):
    yield from ((v, text) for v, text in __whats_new
                if version_cmp(v, last_version) > 0)


def have_new_release(last_version):
    return version_cmp(last_version, version()) < 0
#    for v, text in __whats_new:
#        if version_cmp(v, last_version) > 0:
#            yield (v, text)

#def collect_whats_new(last_version):
#    return [(v, text) for v, text in __whats_new
#            if version_cmp(v, last_version) > 0]
