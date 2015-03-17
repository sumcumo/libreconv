require "libreconv/version"
require "uri"
require "net/http"
require "tmpdir"
require "spoon"


module Libreconv

  def self.convert(source, target, to = "pdf", soffice_command = nil)
    Converter.new(source, target, to, soffice_command).convert
  end

  def self.as_html(source, soffice_command = nil)
    Converter.new(source, '', 'html', soffice_command).as_html
  end

  class Converter
    attr_accessor :soffice_command

    def initialize(source, target, to, soffice_command = nil)
      @source = source
      @target = target
      @to     = to
      @target_path = Dir.tmpdir + "/" + SecureRandom.hex
      @soffice_command = soffice_command
      determine_soffice_command
      check_source_type

      unless @soffice_command && File.exists?(@soffice_command)
        raise IOError, "Can't find Libreoffice or Openoffice executable."
      end
    end

    def convert
      target_tmp_file = _convert
      FileUtils.cp target_tmp_file, @target
    end

    def as_html
      target_tmp_file = _convert
      File.open(target_tmp_file,'rb')
    end

    private

    def _convert
      orig_stdout = $stdout.clone
      $stdout.reopen File.new('/dev/null', 'w')
      pid = Spoon.spawnp(@soffice_command, "--headless", "--convert-to", @to, @source, "--outdir", @target_path)
      Process.waitpid(pid)
      $stdout.reopen orig_stdout
      "#{@target_path}/#{File.basename(@source, ".*")}.#{@to}"
    end

    def determine_soffice_command
      unless @soffice_command
        @soffice_command ||= which("soffice")
        @soffice_command ||= which("soffice.bin")
      end
    end

    def which(cmd)
      exts = ENV['PATHEXT'] ? ENV['PATHEXT'].split(';') : ['']
      ENV['PATH'].split(File::PATH_SEPARATOR).each do |path|
        exts.each do |ext|
          exe = File.join(path, "#{cmd}#{ext}")
          return exe if File.executable? exe
        end
      end

      return nil
    end

    def check_source_type
      is_file = File.exists?(@source) && !File.directory?(@source)
      is_http = URI(@source).scheme == "http" && Net::HTTP.get_response(URI(@source)).is_a?(Net::HTTPSuccess)
      is_https = URI(@source).scheme == "https" && Net::HTTP.get_response(URI(@source)).is_a?(Net::HTTPSuccess)
      raise IOError, "Source (#{@source}) is neither a file nor an URL." unless is_file || is_http || is_https
    end
  end
end
