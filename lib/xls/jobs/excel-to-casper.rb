#
# encoding: utf-8
#
# Copyright (c) 2011-2020 Cloudware S.A. All rights reserved
# Copyright (c) 2011-2020 OCC Ordem dos Contabilistas Certificados. All rights reserved.
#
require 'toconline/loginlib/permissions'
require 'xls/loader'
require 'xls/vrxml/jasper'

class ExcelToCasper
  extend SP::Job::Common

  @@tube_options =  { transient: true }

  def self.tube_options
    @@tube_options
  end

  #
  # JOB execution method main entry point
  #
  def self.perform (job)
    # List of reports that where written by the converter
    replacements = []

    # Make sure only the "special ones" can execute this
    Permissions.validate(job, nil, RoleMask::ROLE_SUPER_USER)

    # copy file from the upload server to the local file system
    update_progress(message: 'A validar excel', progress: 20)
    input_file = get_from_temporary_uploads(file: job[:uploaded_excel], tmp_dir: config[:paths][:temporary_uploads])

    # call the magic sauce
    if job[:converter] == 'vrxml'
      update_progress(message: 'A converter excel para vrxml', progress: 40)
      converter = ::Xls::Vrxml::Converter.new(input_file, true, job[:convert_mode] == 'casper')
      vrxml = converter.convert()
      # find all reports that match the report name all will be overwritten
      report_name  = job[:original_file].split('.')[0]
      replacements = Dir.glob("#{config[:paths][:project]}/app/json_templates/**/#{report_name}.vrxml")
      replacements.each do |replacement|
        File.write(replacement, vrxml)
      end
    else
      update_progress(message: 'A converter excel para jrxml', progress: 40)
      converter = ::Xls::Jrxml::ExcelToJrxml.new(input_file, nil, true, false, job[:convert_mode] == 'casper')

      # find all reports that match the report name all will be overwritten
      report_name  = job[:original_file].split('.')[0]
      replacements = Dir.glob("#{config[:paths][:project]}/app/json_templates/**/#{report_name}.jrxml")
      replacements.each do |replacement|
        File.write(replacement, converter.report.to_xml)
      end
    end

    message = <<-HTML
      <div class="custom-message">
        <casper-icon icon="fa-light:file"></casper-icon>
        <h2 id="title">Conversão completa</h2>
        <div style="display: flex; flex-direction: column; align-items: start;">
    HTML
    if replacements.size != 0
      message += <<-HTML
          <h3 style="margin:0px;">Os seguinte(s) relatório(s) foram <b>temporariamente</b> substituído(s):</h3>
          <ul style="margin:0px;">
            #{to_list(replacements)}
          </ul>
      HTML
    else
      message += <<-HTML
          <h3 style="margin:0px;color: var(--status-red);">Não foi encontrado nenhum relatório para substituir</h3>
      HTML
    end
    message += <<-HTML
        </div>
        <div style="flex-grow: 2.0;"></div>
        <casper-notice type="info" title="Nota">
          Pode retroceder para carregar o EXCEL de novo sem ter que fechar este diálogo.
        </casper-notice>
      </div>
    HTML
    send_response(message: message, custom: true, simple_message: 'Conversão completa')
  end

  def self.to_list (array)
    rv = ''
    array.each do |item|
      rv = rv + "<li style=\"margin:0px;text-align: start;\">#{item.gsub(config[:paths][:project], '.')}</li>"
    end
    rv
  end

end