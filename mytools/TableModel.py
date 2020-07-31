#! /usr/bin/env python
# -*- coding: utf-8 -*-


class TableModel(object):
    def __init__(self, db_name, table_name, field_model_list):
        self.db_name = db_name
        self.table_name = table_name
        self.field_model_list = field_model_list


class FieldModel(object):
    def __init__(self, column_name, column_type, data_type, character_maximum_length, is_nullable, column_default,
                 column_comment):
        self.column_name = column_name
        self.column_type = column_type
        self.data_type = data_type
        self.character_maximum_length = character_maximum_length
        self.is_nullable = is_nullable
        self.column_default = column_default
        self.column_comment = column_comment

    def print_model(self):
        print('%s: %s' % (self.column_name, self.column_type))
