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
        'parameter': Parameter.capture,
        'field': Field.capture,
        'variable': Variable.capture
      }

      #
      # Extract parameter(s)/field(s)/variable(s) from an expression.
      #
      # @param expression: V8 expression
      #
      def self.extract(expression:, caller: caller_locations(1,1)[0].base_label)
        all = []
        done = []
        # collect
        PFV_EXPR.each do | key, exp |
          expression.scan(exp) { | v |
            if false == done.include?(v[0])
              all << { type: key, value: v[0] }
              done << v[0]
            end
          }
        end
        # log?
        if all.count > 0 && ( Vrxml::Log::EXTRACTION == ( Vrxml::Log::MASK & Vrxml::Log::EXTRACTION ) )
          print "#{caller} called Expression.#{__method__}: ".cyan
          print "#{expression}".yellow
          print " #{all.count > 1 ? 'expression' : 'param/field/variable' }\n".cyan
          ap all
        end
        # done
        all
      end

      #
      # Translate a JRXML JAVA expression to VRXML expression.
      #
      # @param expression JRXML JAVA expression to translate.
      #
      # @return VRXML expression
      #
      def self.translate(expression:, relationship:, nce:, tracking: { file: __FILE__, line: __LINE__, method: __method__, caller: caller_locations(1,1)[0].base_label })
        _exp = expression.strip
        _ext = []
        if 0 == _exp.length
          return _exp, _ext
        end
        _exp = Vrxml::Expression.i_translate(expression: _exp, relationship: relationship, nce: nce, tracking: tracking)
        ( Vrxml::Expression.extract(expression: _exp) || [] ).each do | e |
          case e[:type]
          when :parameter, :field, :variable
            _ext << e
          else
            raise "???"
          end # case
        end # each
        # log
        ::Xls::Vrxml::Log.TRANSLATION(from: expression, to: _exp, tracking: tracking)
        # done
        return _exp, _ext
      end

      private

      #
      # Translate a JAVA expression into a JS ( V8 ) expression.
      #
      # @param uri JRXML local URI
      # @param expression
      # @param relationship
      # @param nc 		  Not converted expressions.
      #

      LEGACY_PARAM_EXP = /\$[P]\{([a-zA-Z0-9_]+)\}/
      LEGACY_FIELD_EXP = /\$[F]\{([a-zA-Z0-9_]+)\}/
      LEGACY_VAR_EXP   = /\$[V]\{([a-zA-Z0-9_]+)\}/
      LEGACY_EXP       = { parameter: LEGACY_PARAM_EXP, field: LEGACY_FIELD_EXP, variable: LEGACY_VAR_EXP }

      def self.i_translate(uri: 'TODO', expression:, relationship:, nce:, tracking:)
        # already translated?
        if expression.include?("$.") || expression.include?("$['")
          # ... done ...
          return expression
        end
        # not yet, translate it now
        error = false
        Open3.popen3('jrxml2vrxml', '-s', "#{expression}", '-r', "#{relationship || '<replace-me>'}") do |stdin, stdout, stderr, wait_thr|
          if 0 != ( wait_thr.value.to_i >> 8 )
            # log
            ::Xls::Vrxml::Log.WARNING(msg: "Unable to convert expression: #{expression}")
            # track
            nce[expression] = {
              :jrxml => uri,
              :error => JSON.parse(stderr.read.strip, symbolize_names: true)[:error]
            }
            error = true
            # TODO 2.0? is this ok ?
            return expression
          end
          rv = JSON.parse(stdout.read.strip, symbolize_names: true)[:translated]
          # no param/field/variable
          if 0 == rv.length
            rv = expression
          end                    
          # done
          return rv
        end # popen3
      end      

    end # class 'Expression'

  end # of module 'Vrxml'
end # of module 'Xls'
