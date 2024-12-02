# frozen_string_literal: true

source "https://rubygems.org"

# Specify your gem's dependencies in acrobat.gemspec
gemspec

group :development do
  eval_gemfile "gemfiles/rubocop.gemfile"
end

gem "guard", "~> 2.19", groups: %i[development test]
gem "guard-minitest", "~> 2.4", groups: %i[development test]
