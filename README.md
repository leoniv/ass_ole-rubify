# AssOle::Rubify

It's a part of [ass_ole](https://github.com/leoniv/ass_ole) stack. Gem
`ass_ole-rubify` provides wrappers which make `Ruby` scripting
over 1C `WIN32OLE` objects easly and shortly as the traditional `Ruby`
scripting.

All wrappers is kinde of `AssOle::Rubify::GenericWrapper` and always holds two
things. First thing `ole` is a wrapped `WIN32OLE` object. Second thing
`ole_runtime` is a 1C ole connector which spawned the `ole` object
(`ole_runtime` recognize `ole` as native 1C object but not a `ComObject`).

`GenericWrapper` provides generic API and all missing method sends to `ole` in
`method_missing` handler. Custom wrappers provides custom API depended
on type of `ole`.

All custom wrappers, can includes various mixins for dynamic generated
wrapper API and always yields instance in end of constructor where instance
can be extended would you like.

FIXME:
>  Custom wrappers included into `AssOle::Rubify::SchemaObjects` namespace
>  wrapps objects defined in 1C application *Meta Data Tree*.
>  In 1C term sach objects named as *Applied Objects*.
>  All wrappers from `SchemaObjects` holds therd thing `md_manger` instance wich
>  wrapp *Objects manager* like *Catalogs.CatalogName*, *Documents.DocumentName*
>  etc. Wrappers for `md_manger` dynamicaly generated in
>  `AssOle::Rubify::MdManagers` namespace for specific 1C application instance.

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

`AssOle::Rubify` provides wrappers for does the same action above in the Ruby
style

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
