# test/win32ole/jruby/test_array.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/array'

class TestSafeArray < Test::Unit::TestCase
  SA = WIN32OLE::SafeArray

  def test_dimension_count_for_a_flat_array
    assert_equal(1, SA.dimension_count([1, 2, 3]))
  end

  def test_dimension_count_for_a_nested_array
    assert_equal(2, SA.dimension_count([[1, 2], [3, 4]]))
  end

  def test_dimension_count_takes_the_max_depth_across_branches
    assert_equal(2, SA.dimension_count([[1, 2], 3]))
  end

  def test_dimension_sizes_for_a_2d_array
    assert_equal([2, 3], SA.dimension_sizes([[1, 2, 3], [4, 5, 6]]))
  end

  def test_dimension_sizes_takes_the_max_across_ragged_branches
    assert_equal([2, 3], SA.dimension_sizes([[1, 2, 3], [4]]))
  end

  def test_nested_entry_reads_the_addressed_element
    ary = [[1, 2], [3, 4]]
    assert_equal(4, SA.nested_entry(ary, [1, 1]))
  end

  def test_nested_entry_returns_nil_past_a_shorter_branch
    ary = [[1, 2, 3], [4]]
    assert_nil(SA.nested_entry(ary, [1, 2]))
  end

  def test_each_fill_index_enumerates_the_innermost_dimension_fastest
    tuples = SA.each_fill_index([2, 3]).to_a
    assert_equal([[0, 0], [0, 1], [0, 2], [1, 0], [1, 1], [1, 2]], tuples)
  end

  def test_each_fill_index_round_trips_a_2d_array_through_nested_entry
    ary = [[1, 2, 3], [4, 5, 6]]
    sizes = SA.dimension_sizes(ary)
    values = SA.each_fill_index(sizes).map { |pid| SA.nested_entry(ary, pid) }
    assert_equal([1, 2, 3, 4, 5, 6], values)
  end

  def test_each_read_index_enumerates_the_outermost_dimension_fastest
    tuples = SA.each_read_index([0, 0], [1, 2]).to_a
    assert_equal([[0, 0], [1, 0], [0, 1], [1, 1], [0, 2], [1, 2]], tuples)
  end

  def test_each_read_index_respects_nonzero_lower_bounds
    tuples = SA.each_read_index([1, 1], [2, 2]).to_a
    assert_equal([[1, 1], [2, 1], [1, 2], [2, 2]], tuples)
  end
end
end
