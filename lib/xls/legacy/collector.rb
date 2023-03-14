# encoding: utf-8
#
# Copyright (c) 2011-2023 Cloudware S.A. All rights reserved.
#
# This file is part of xls2vrxml.
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

require_relative 'bands'
require_relative 'binding'

module Xls
  module Legacy

    class Collector
      
      attr_accessor :bands
      attr_accessor :binding

      #
      # Initialize a 'collector'.
      #
      # @param layout       'Layout' sheet
      # @param binding      'Data Binding' sheet
      # @param relationship for translation purpose.
      # @param hammer       no comments
      #
      def initialize(layout:, binding:, relationship:'lines', hammer: nil)
        @bands   = Bands.new(sheet: layout, relationship: relationship, hammer: hammer)
        @binding = Binding.new(sheet: binding, relationship: relationship, hammer: hammer)
      end

    end # of class 'Collector'

  end # of module 'Legacy'
end # of module 'Xls'