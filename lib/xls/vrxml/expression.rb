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
require_relative 'object'
require_relative 'log'

module Xls
  module Vrxml

    class Expression < Object

      PFV_EXPR = {
        'param': Parameter.expr,
        'field': Field.expr,
        'variable': Variable.expr
      }

      #
      # Extract parameter(s)/field(s)/variable(s) from an expression.
      #
      # @param expression: V8 expression
      #
      def self.extract(expression:, caller: caller_locations(1,1)[0].base_label)
        all = []
        # collect
        PFV_EXPR.each do | key, exp |
          expression.scan(exp) { | v |
            all << { type: key, value: v }
          }
        end
        # log?
        if all.count > 0 && ( Vrxml::Log::EXTRACTION == ( Vrxml::Log::FLAGS & Vrxml::Log::EXTRACTION ) )
          print "#{caller} called Expression.#{__method__}: ".cyan
          print "#{expression}".yellow
          print " #{all.count > 1 ? 'expression' : 'param/field/variable' }\n".cyan
          ap all
        end
        # done
        all
      end

      #
      # Translate a JAVA expression into a JS ( V8 ) expression.
      #
      # @param uri JRXML local URI
      # @param expression
      # @param relationship
      # @param nc 		  Not converted expressions.
      #
      def self.translate(uri:, expression:, relationship:, nc:, caller: caller_locations(1,1)[0].base_label)
        # already translated?
        if expression.include?("$.") || expression.include?("$['")
          # ... done ...
          return expression
        end
        # not yet, translate it now
        error = false
        Open3.popen3('jrxml2vrxml', '-s', "#{expression}", '-r', "#{relationship || '<replace-me>'}") do |stdin, stdout, stderr, wait_thr|
          if 0 != ( wait_thr.value.to_i >> 8 )
            puts "UNABLE TO CONVERT: #{expression}".red
            nc[expression] = {
              :jrxml => uri,
              :error => stderr.read
            }
            error = true
            return expression
          end
          rv = stdout.read.strip
          # no param/field/variable
          if 0 == rv.length
              rv = expression
          end
          # log?
          if expression != rv && ( Vrxml::Log::TRANSLATION == ( Vrxml::Log::FLAGS & Vrxml::Log::TRANSLATION ) )
              puts "#{caller} called Expression.#{__method__}:".cyan
              puts "  '%s' ~> '%s'" %[ "#{expression}".yellow, "#{rv}".green ]
          end
          # success
          return rv
        end # popen3
      end

    end # class 'Expression'

  end # of module 'Vrxml'
end # of module 'Xls'
