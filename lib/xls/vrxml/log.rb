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
  module Vrxml

    class Log

      SHOW_INFO       = 0x0001
      SHOW_STATS      = 0x0002
      SHOW_PROPERTIES = 0x0004
      TABLES          = 0x0008
      EXTRACTION      = 0x0010
      TRANSLATIONS    = 0x0020
      STEPS           = 0x0040
      WARNINGS        = 0x0080
      ERRORS          = 0x0100
      TRACE           = 0x0200
      TODO            = 0x0400
      WHAT_IS         = 0x0800
      DEBUG           = 0x1000
      SHOW_V8_EXPR    = 0x0FFF

      MASK           = SHOW_INFO | STEPS | WARNINGS | ERRORS | TODO | DEBUG # | TRANSLATIONS
      # | TABLES
      # | TRACE | TRANSLATIONS

      public

      def self.LOG_IF(msg:, bits:)
        if ( bits == ( Vrxml::Log::MASK & bits ) )
          puts "%s" % [ msg ]
        end
      end

      def self.TODO(msg:, caller: caller_locations(1,1)[0].base_label)
        if ( Vrxml::Log::TODO == ( Vrxml::Log::MASK & Vrxml::Log::TODO ) )
          puts "⌄".purple
          puts " TODO 2.0: %s".purple % [ "#{msg}".white ]
          puts "⌃".purple
        end
      end

      def self.WHAT_IS(msg:, caller: caller_locations(1,1)[0].base_label)
        if ( Vrxml::Log::WHAT_IS == ( Vrxml::Log::MASK & Vrxml::Log::WHAT_IS ) )
          puts "⌄".purple
          puts " WHAT IS: %s".purple % [ "#{msg}".white ]
          puts "⌃".purple
        end
      end

      def self.TRACE(who:, what:)
        if ( Vrxml::Log::TRACE == ( Vrxml::Log::MASK & Vrxml::Log::TRACE ) )
          puts " ⌁ %s called %s".purple % [ who, what ]
        end
      end

      def self.TRANSLATION(from:, to:, tracking: nil)
        if ( Vrxml::Log::TRANSLATIONS == ( Vrxml::Log::MASK & Vrxml::Log::TRANSLATIONS ) )
          puts " ➢ TRANSLATION ( @ %s ) :\n    %s \n    %s".white % [ "#{tracking[:caller]}".yellow, "#{from}".cyan, "#{to}".green ]
        end
      end

      def self.WARNING(msg:)
        if ( Vrxml::Log::WARNINGS == ( Vrxml::Log::MASK & Vrxml::Log::WARNINGS ) )
          puts " ⚠︎ %s".yellow % [ msg ]
        end
      end

      def self.ERROR(msg:, exception: nil)
        if ( Vrxml::Log::ERRORS == ( Vrxml::Log::MASK & Vrxml::Log::ERRORS ) )
          puts " ⨯ %s".red % [ msg ]
        end
        if nil != exception
          raise exception, msg
        end
      end

      def self.STEP(msg:)
        if ( Vrxml::Log::STEPS == ( Vrxml::Log::MASK & Vrxml::Log::STEPS ) )
          puts "#{msg}"
        end
      end

    end # class 'Log'

  end # of module 'Vrxml'
end # of module 'Xls'
