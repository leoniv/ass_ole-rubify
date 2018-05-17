# AssOle::Rubify

It's a part of [ass_ole](https://github.com/leoniv/ass_ole) stack. Gem
`ass_ole-rubify` provides wrappers which make `Ruby` scripting
over 1C `WIN32OLE` objects easly and shortly as the traditional `Ruby`
scripting.

General class in this gem is the `GenericWrapper`. `GenericWrapper` always holds
two things. First thing is the wrapped `WIN32OLE` object. Second thing is the
1C ole runtime (ComConnector) which spawned the `WIN32OLE` object.

`GenericWrapper` sends all missing method to `WIN32OLE` object in the
`method_missing` handler.

Next main thing in this gem is the `GenericWrapper::Mixins`. Mixins provides
methods for "rubify" `WIN32OLE` objects wrapped in the `GenericWrapper`.

All modules from `GenericWrapper::Mixins` namespace automatically includes in
the `GenericWrapper` instances depending on the `WIN32OLE` object type.

Third thing of this gem is the `GlobContext` - special wrapper for
"global context" methods like a `NewObject` etc for wrapping all `WIN32OLE`
in the `GenericWrapper`.

`GlobContext` also has `GlobContext::Mixins` which includes in
the `GlobContext` depending on the ole runtime type. `GlobContext::Mixins` isn't
provides any methods and defined for customize they as you need in your
application. Except only `GlobContext::Mixins::ServerContext` provides `#query`
wrapper.

## About `GenericWrapper::Mixins`

TODO: write this

## General usage

You can usage `GenericWrapper` as standalone class and wrapping selected
`WIN32OLE` objects manually with `Rubify.rubify` or `Rubify#rubify` methods.

Also you can usage `GlobContext` wrapper which automatically wrapping all
`WIN32OLE` in `GenericWrapper`. For it you should define your class or module
as `like_rubify_runtime YourOleRuntime`:
```
class Foo
  like_rubify_runtime Runtimes::External
  ...
end
```

And last general usage case is to write own wrapper class, child of the
`GenericWrapper` and own wrapper mixins. But this case needs a some
explanations.

TODO: write about make custom wrapper class.


## Main profit from Rubify

Standatrt scripting ower 1C with [ass_ole runtime](https://github.com/leoniv/ass_ole)

```ruby
require 'ass_ole'
require 'ass_maintainer-info_base'

INFO_BASE = AssMaintainer::InfoBase.new('', 'File="path"')

module ExtRuntime
  is_ole_runtime :external
  run INFO_BASE
end

module Worker
  like_ole_runtime ExtRuntime
end


# Working with Catalogs
cat_item = Worker.Catalogs.CatalogName.CreateItem #=> WIN32OLE
cat_item.Description = 'Item name'
cat_item.Attr1 = 'Attr 1 value'

3.times do |i|
  row = cat_item.TabularSection1.Add
  row.Attr1 = "Attr1 val #{i}"
  row.Attr2 = "Attr2 val #{i}"
end

cat_item.Write #=> nil

# Workin with 1C Array
ole_arr = Worker.newObject('Array') #=> WIN32OLE

3.times do |i|
  ole_arr.Add(i)
end

# Maping arr
new_arr = []

ole_arr.Count.times do |index|
  new_arr << arr.Get(index) ** 2
end

# Select from Array
new_arr = []

ole_arr.Count.times do |index|
  val = ole_arr.Get(index)
  new_arr << val if val > 1
end

# ... etc

```

`AssOle::Rubify` provides wrapper and its mixins for does the same action
above in the Ruby style

```ruby
require 'ass_ole-rubify'
require 'ass_maintainer-info_base'

INFO_BASE = AssMaintainer::InfoBase.new('', 'File="path"')

module ExtRuntime
  is_ole_runtime :external
  run INFO_BASE
end

module Worker
  like_rubify_runtime ExtRuntime
end

# Working with Catalogs
cat_item = Worker.Catalogs.CatalogName
  .CreateItem(Description: 'Item name', Attr1: 'Attr1 value') do |item|

  3.times do |i|
    item.TabularSection1.Add Attr1: "Attr1 val #{i}" do |row|
      row.Attr2 = "Attr2 val #{i}"
    end
  end
# Not `Write`! Invoke `write` method defined in GenericWrapper::Mixins::Write
end.write #=> AssOle::Rubify::GenericWrapper

# Workin with 1C Array
arr = Worker.newObject('Array') do |a|
  3.times do |i|
    a.Add(i)
  end
end #=> AssOle::Rubify::GenericWrapper

# Maping arr
arr.map do |item|
  item ** 2
end #=> Array

# Select from arr
arr.select do |item|
  item > 1
end #=> Array

# ...etc
```


## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ass_ole-rubify'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ass_ole-rubify

## Usage

TODO: Write usage instructions here

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/ass_ole-rubify.
