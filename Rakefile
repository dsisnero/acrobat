# -*- ruby -*-

require "rubygems"
require "hoe"

class Hoe

  def intuit_values
    require 'asciidoctor'
    doc = Asciidoctor.load_file readme_file
    urls = { "home" => doc.attributes['uri-repo']}
    desc = doc.blocks.find{|b| b.id =~ /description/i}.content
    summ = desc.split(/\.\s+/).first(summary_sentences).join(". ")
    self.urls ||= urls
    self.description ||= desc
    self.summary ||= summ
  end

end

# Hoe.plugin :compiler
# Hoe.plugin :gem_prelude_sucks
# Hoe.plugin :inline
# Hoe.plugin :racc
# Hoe.plugin :rcov
# Hoe.plugin :rdoc
Hoe.plugin :bundler
Hoe.plugin :minitest
Hoe.plugin :yard

Hoe.spec "acrobat" do |s|
  dependency("hoe-bundler", "~> 1.4",:development)
  dependency("hoe-yard", "> 0.1",:development)
  dependency("pry", "> 0.0", :development)
  dependency("pry-byebug", "> 0.0", :development)
  dependency("yard", "> 0.0", :development)
  dependency("guard", "> 0.0", :development)
  dependency("wdm", "> 0.1", :development) if Gem.win_platform?
  dependency("guard-minitest", "> 0.0", :development)
  dependency("asciidoctor", "> 0.0",:development)
#  dependency("minitest-utils", "> 0.0", :development)
  developer("Dominic Sisneros","dsisnero@gmail.com")
  clean_globs << "tmp"
  license "MIT" # this should match the license in the README
end


# vim: syntax=ruby
