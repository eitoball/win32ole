require "test/unit"
require "core_assertions"

Test::Unit::TestCase.include Test::Unit::CoreAssertions

if RUBY_ENGINE == 'jruby'
  # Unlike CRuby's (a full copy of RbConfig::CONFIG), JRuby's
  # RbConfig::MAKEFILE_CONFIG is a small, hand-picked subset that omits
  # RUBY_INSTALL_NAME -- mkmf.rb's `CONFIG = RbConfig::MAKEFILE_CONFIG`
  # then leaves CONFIG["RUBY_INSTALL_NAME"] nil wherever a test builds a
  # path from it (e.g. test/win32ole/test_err_in_callback.rb's setup).
  # Same object as mkmf's CONFIG (not a dup), so this backfill is visible
  # there regardless of require order.
  RbConfig::MAKEFILE_CONFIG['RUBY_INSTALL_NAME'] ||= RbConfig::CONFIG['RUBY_INSTALL_NAME']
end
