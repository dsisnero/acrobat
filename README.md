# DESCRIPTION:

The acrobat gem is a library that uses WIN32OLE to automate pdf
functions. Adobe Reader or Adobe Acrobat must be installed and then you
can fill forms and merge documents

# Requirements

-   Windows Adobe Acrobat installed

# Install 0ptions

acrobat can be installed using (a) the `gem install` command, (b)
Bundler

## (a) gem install

Open a terminal and type \$ gem install acrobat

:::: tip
::: title
Upgrading your installation
:::

If you have an earlier version of acrobat installed you can update it
using:

\$ gem update acrobat

If you install a new version of the gem using `gem install` instead of
gem update, you'll have multiple versions installed. If that's the case,
use the following gem command to remove the old versions:

\$ gem cleanup acrobat
::::

## (b) Bundler

1.  Create a Gemfile in the root folder of your project (or the current
    directory)

2.  Add the `asciidoctor` gem to your Gemfile as follows:

    ``` ruby
    source 'https://rubygems.org'
    gem 'acrobat'
    # or specify the version explicitly
    # gem 'acrobat', '0.0.5'
    ```

3.  Save the Gemfile

4.  Open a terminal and install the gem using:

        $ bundle

To upgrade the gem, specify the new version in the Gemfile and run
`bundle` again. Using `bundle update` is **not** recommended as it will
also update other gems, which may not be the desired result.

# Usage

``` ruby
  require 'acrobat'
  Adobe::Acrobat.run do |app|
    fields = {city: 'City', state: 'State', zip: 'zip'}
    doc = app.open('apd_document.pdf')
    doc.fill_form(fields)
    doc.show
    doc2 = app.open('another_doc.pdf')
    doc2.fill_form(fields)
    doc1.merge_doc(doc2)
    doc1.save_as(name: 'merged_doc.pdf', dir: 'tmp')
  end
```

# Getting Help

We encourage you to ask questions and discuss any aspects of the project
on the discussion list, Twitter or IRC.

Mailing list

:   {uri-discuss}

The project's source code, issue tracker, and sub-projects are on github

Source repository (git)

:   <https://github.com/dsisnero/acrobat>

Issue tracker

:   <https://github.com/dsisnero/acrobat/issues>

acrobat organization on GitHub

:   <https://github.com/dsisnero>

# Copyright and Licensing

Copyright © 2012-2016 Dominic Sisneros and the acrobat-Ruby Project.
Free use of this software is granted under the terms of the MIT License.

See the
[LICENSE](https://github.com/dsisnero/acrobat/blob/master/LICENSE.txt)
file for details.
