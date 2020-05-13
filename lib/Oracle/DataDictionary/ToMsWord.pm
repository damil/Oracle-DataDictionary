package Oracle::DataDictionary::ToMsWord;
use utf8;
use Moose;
use Template;
use MsOffice::Word::HTML::Writer;
use Path::Tiny qw/path tempfile/;
use Data::Reach     qw/reach/;
use Clone           qw/clone/;

our $VERSION = 0.1;

#======================================================================
# ATTRIBUTES AND BUILDER METHODS
#======================================================================

has 'dict'        => (is => 'ro', isa => 'Oracle::DataDictionary', required   => 1);
has 'config_data' => (is => 'ro', isa => 'HashRef', required => 1);
                    # TODO : check config domain

has 'name'        => (is => 'ro', isa => 'Str', lazy_build => 1);
has 'd_extract'   => (is => 'ro', isa => 'Str', default    => scalar(localtime));

has 'tmpl_path'   => (is => 'ro', isa => 'Str', lazy_build => 1);
has 'template'    => (is => 'ro', isa => 'Str', default    => 'data_dictionary.tt');


has 'tablegroups' => (is => 'ro', isa => 'ArrayRef', lazy_build => 1, init_arg => undef);


sub _build_name {my $self = shift; $self->dict->owner}

sub _build_tmpl_path {
  my $self = shift;
  my $class = ref $self;

  # find a "share" directory under the location of the module
  $class =~ s[::][/]g;
  my $path = $INC{$class . ".pm"};
  $path =~ s[\.pm$][/share];
  return $path;
}


sub _build_tablegroups {
  my ($self) = @_;

  my $tables = clone $self->dict->tables;

  # merge with descriptions from config
  foreach my $table_name (keys %$tables) {
    $tables->{$table_name}{colgroups} = $self->_colgroups($table_name);

    my $descr = $self->_config(tables => $table_name => 'descr');
    $tables->{$table_name}{descr} = $descr if $descr;
  }

  # grouping: merge with table info from config
  my $tablegroups = clone $self->_config('tablegroups') || [];


  foreach my $group (@$tablegroups) {
    # tables declared in this group are removed from the global %$tables ..
    my @declared_table_names = @{$group->{tables}};
    my @extracted_tables     = grep {$_} map {delete $tables->{$_}} @declared_table_names;

    # .. and their full definitions take place of the declared names
    $group->{tables} = \@extracted_tables;
  }

  # deal with remaining tables (
  if (my @other_tables = sort keys %$tables) {

    # Filter out based on the regexps in filters include & exclude
    if (my $filter_include = $self->_config(qw/filters include/)) {
      @other_tables = grep { $_ =~ /$filter_include/ } @other_tables;
    }
    if (my $filter_exclude = $self->_config(qw/filters exclude/)) {
      @other_tables = grep { $_ !~ /$filter_exclude/ } @other_tables;
    }

    # if some unclassified tables remain after the filtering
    if (@other_tables) {
      push @$tablegroups, {
        name   => 'Unclassified tables', 
        descr  => 'Present in database but unlisted in config',
        tables => [ @{$tables}{@other_tables} ],
      };
    }
  }

  return $tablegroups;
}


#======================================================================
# PUBLIC METHODS
#======================================================================

sub render {
  my ($self, $target) = @_;

  my $data = {name        => $self->name,
              tablegroups => $self->tablegroups,
              d_extract   => $self->d_extract};

  # Template Toolkit renderer
  my $renderer = Template->new(
    LOAD_PERL    => 1, # so that the template can call MsOffice::Word::HTML::Writer
    INCLUDE_PATH => $self->tmpl_path,
   );

  # document generation; beware the unusual pattern : result will be in
  # $target, so $output should be empty
  $renderer->process($self->template,
                     {data => $data, target => $target},
                     \my $output)
    or die $renderer->error;

  # check that $output is really empty
  die $output if $output && $output !~ /^\s+$/;
}



#======================================================================
# PRIVATE METHODS
#======================================================================


sub _config {
  my ($self, @path) = @_;
  return reach $self->config_data, @path;
}


sub _colgroups {
  my ($self, $table_name) = @_;

  # columns
  my $db_columns = clone $self->dict->tables->{$table_name}{col} // {};

  # grouping: merge with column info from config
  my $colgroups = clone $self->_config(tables => $table_name => 'colgroups') || [];
  foreach my $group (@$colgroups) {
    my @columns;
    foreach my $column (@{$group->{columns}}) {
      my $col_name = $column->{name};
      my $db_col = delete $db_columns->{$col_name} or next;
      push @columns, {%$db_col, %$column};
    }
    $group->{columns} = \@columns;
  }

  # deal with remaining columns (present in database but unlisted in
  # config); sorted with primary keys first, then alphabetically.
  my $sort_pk = sub {   $db_columns->{$a}{is_prim_key} ? -1
                      : $db_columns->{$b}{is_prim_key} ?  1
                      :                                  $a cmp $b};
  if (my @other_cols = sort $sort_pk keys %$db_columns) {
    # build colgroup
    push @$colgroups, {name    => 'Unclassified columns', 
                       columns => [ @{$db_columns}{@other_cols} ]};
  }

  return $colgroups;
}





1;

__END__

=head1 NAME

Oracle::DataDictionary::ToMsWord

=head1 SYNOPSIS

  use DBI;
  use Oracle::DataDictionary;
  use Oracle::DataDictionary::ToMsWord;n

  my $dbh = DBI->connect("dbi:Oracle:...", ...);
  my dict = Oracle::DataDictionary->new(dbh => $dbh);
  my $msw = Oracle::DataDictionary::ToMsWord->new(
    dict        => $dict,
    config_data => $config,
   );
  $msw->render($path_to_msword);

  # then open the doc in Microsoft Word, select all and click F9 to refresh
  # the table of contents and column index


=head1 DESCRIPTION

This module produces a Microsoft Word file that documents
an Oracle schema. The document content is a mix of technical data
coming from Oracle data dictionary tables and contextual data coming
from a configuration tree, that can specify :

=over

=item *

grouping and ordering information for tables;

=item *

grouping and ordering information for columns;

=item *

textual descriptions of the purpose of tables and columns;

=item *

some presentation options (for example some technical columns
may be presented in an abbreviated form).

=back


=head1 METHODS

=head2 new

  my $msw = Oracle::DataDictionary::ToMsWord->new(
    dict        => $dict,
    config_data => $config,
   );

Creates an instance of a MsWord renderer. Arguments to the C<new()> method are:

=over

=item dict

mandatory. Reference to an L<Oracle::DataDictionary> object

=item config_data

mandatory. Hashref to a datatree that matches the structure described below
in the L</CONFIG> section

=item name

optional. Name that will be displayed in the MsWord title. Defaults to the 
owner of the Oracle schema associated with the C<$dict> object.

=item d_extract

optional. Date of extraction, to be displayed in the MsWord title.
Defaults to the current date.


=item tmpl_path

optional. Path to a directory where L<Template> Toolkit looks for templates.
Defaults to a "share" directory under the current module.

=item template

optional. Name of the rendering template for the L<Template> Toolkit. 
Defaults to C<data_dictionary.tt>.

=back

=head2 render

  $msw->render($path_to_msword);

Produces a MsWord document stored at the specified location.

The document is in C<.mht> format, not C<.docx>.

It is strongly advised to open the file in MsWord, select all contents
by pressing C<Ctrl-a>, refresh all fields by pressing C<F9>, and store
the result in C<.docx> format.

While refreshing fields, MsWord will automatically generate the table
of contents at the beginning of the document, and the column index at the
end of the document.

=head1 CONFIGURATION TREE

The configuration tree can be specified in pure Perl, or it can
be read from a YAML, JSON or XML file.
The data structure is the same as used by L<App::AutoCRUD>;
therefore it is possible to use the same config information both
for generating documentation and for running an app that displays
or modifies data from the Oracle database.

The expected datas tructure is as described in 
L<App::AutoCRUD::ConfigDomain/datasources>.


[TODO : Chinook example]


=head1 AUTHOR

Laurent Dami, C<< <laurent dot dami at cpan dot org> >>, May 2020.


=head1 LICENSE AND COPYRIGHT

Copyright 2020 Laurent Dami.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

