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
