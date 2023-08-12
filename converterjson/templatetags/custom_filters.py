import os

from django import template

register = template.Library()


@register.filter
def basename(value):
    return os.path.basename(value)


@register.filter(name='split_filename')
def split_filename(value):
    return value.split('/')[-1]
