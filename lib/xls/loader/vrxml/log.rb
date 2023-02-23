# encoding: utf-8
#
# Copyright (c) 2011-2023 Cloudware S.A. All rights reserved.
#
# This file is part of xls2vrxml.
# Based on sp-excel-loader.
#
# xls2vrxml is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# xls2vrxml is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with xls2vrxml.  If not, see <http://www.gnu.org/licenses/>.
#

require 'open3'

module Xls
  module Loader
    module Vrxml

      class Log

        SHOW_INFO       = 0x01
        SHOW_STATS      = 0x02
        SHOW_PROPERTIES = 0x04
        TABLES          = 0x08
        EXTRACTION      = 0x10
        TRANSLATION     = 0x20
        SHOW_V8_EXPR    = 0x7F

        FLAGS           = SHOW_INFO

      public

      def self.TODO(msg:, caller: caller_locations(1,1)[0].base_label)
        puts "⌄".purple
        puts " TODO 2.0: %s".purple % [ "#{msg}".white ]
        puts "⌃".purple
      end

      def self.WHAT_IS(msg:, caller: caller_locations(1,1)[0].base_label)
        puts "⌄".purple
        puts " WHAT IS: %s".purple % [ "#{msg}".white ]
        puts "⌃".purple
      end

      end # class 'Log'

    end # module 'Vrxml'
  end # module 'Loader'
end # module 'Xls'
