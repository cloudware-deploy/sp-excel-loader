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
# encoding: utf-8
#
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xls/loader/version'

Gem::Specification.new do |spec|
  spec.name          = 'xls2vrxml'
  spec.version       = Xls::Loader::VERSION
  spec.email         = ['emi@cldware.com']
  spec.bindir        = 'bin'
  spec.executables   = ['xls2jrxml']
  spec.date          = '2012-10-17'
  spec.summary       = 'xls2vrxml'
  spec.description   = 'Extends RubyXL adding handling of excel tables and other conversion utilies'
  spec.authors       = ['Cloudware S.A.']
  spec.files         = Dir.glob("lib/**/*") + Dir.glob("spec/**/*") + %w(LICENSE README.md Gemfile)
  spec.homepage      = 'https://github.com/cloudware-deploy/xls2vrxml'
  spec.license       = 'AGPL 3.0'
  spec.require_paths = ['lib']

  spec.add_development_dependency 'bundler'
  spec.add_development_dependency 'rspec'
  spec.add_development_dependency 'byebug'

  spec.add_dependency 'awesome_print'
  spec.add_dependency 'rubyXL'
  spec.add_dependency 'json'
end
