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

module Xls
  module Loader
    module Jrxml

      class Property

        attr_accessor :name
        attr_accessor :value

        def initialize (a_name, a_value)
          @name  = a_name
          @value = a_value
        end

        def attributes
          { name: @name, value:@value }
        end

        def to_xml(a_node)
          Nokogiri::XML::Builder.with(a_node) do |xml|
            xml.property(attributes)
          end
        end

      end

    end

  end
end
