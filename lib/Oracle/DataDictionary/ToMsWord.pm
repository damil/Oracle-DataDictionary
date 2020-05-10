package Oracle::DataDictionary::ToMsWord;
use utf8;
use Moose;
use Template;
use Template::Filters ();
use Encode qw/encode_utf8/;
use MsOffice::Word::HTML::Writer;
use Path::Tiny qw/path tempfile/;
use Data::Reach     qw/reach/;
use Clone           qw/clone/;


has 'dict'        => (is => 'ro', isa => 'Oracle::DataDictionary', required   => 1);
has 'config_data' => (is => 'ro', isa => 'HashRef', required => 1);

# TODO : check config domain

has 'd_extract'   => (is => 'ro', isa => 'Str', default => scalar(localtime));
has 'name'        => (is => 'ro', isa => 'Str', lazy_build => 1);

has 'tablegroups' => (is => 'ro', isa => 'ArrayRef', lazy_build => 1);
has 'template'    => (is => 'ro', isa => 'Str', default => 'data_dictionary.tt');
has 'tmpl_path'   => (is => 'ro',  isa => 'Str', lazy_build => 1);




sub _build_name {my $self = shift; $self->dict->owner}

sub _build_tmpl_path {
  my $self = shift;
  my $class = ref $self;

  $class =~ s[::][/]g;
  my $path = $INC{$class . ".pm"};
  $path =~ s[\.pm$][/share];
  return $path;
}



sub config {
  my ($self, @path) = @_;
  return reach $self->config_data, @path;
}





sub _build_tablegroups {
  my ($self) = @_;

  my $tables = clone $self->dict->tables;

  # merge with descriptions from config
  foreach my $table_name (keys %$tables) {
    $tables->{$table_name}{colgroups} = $self->colgroups($table_name);

    my $descr = $self->config(tables => $table_name => 'descr');
    $tables->{$table_name}{descr} = $descr if $descr;
  }

  # grouping: merge with table info from config
  my $tablegroups = clone $self->config('tablegroups') || [];


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
    if (my $filter_include = $self->config(qw/filters include/)) {
      @other_tables = grep { $_ =~ /$filter_include/ } @other_tables;
    }
    if (my $filter_exclude = $self->config(qw/filters exclude/)) {
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



sub colgroups {
  my ($self, $table_name) = @_;

  # columns
  my $db_columns = clone $self->dict->tables->{$table_name}{col} // {};

  # grouping: merge with column info from config
  my $colgroups = clone $self->config(tables => $table_name => 'colgroups') || [];
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

  # document generation
  $renderer->process($self->template,
                     {data => $data, target => $target},
                     \my $output)
    or die $renderer->error;

  die $output if $output && $output !~ /^\s+$/;
}



1;

__END__

