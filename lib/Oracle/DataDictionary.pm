package Oracle::DataDictionary;
use utf8;
use Moose;
use Moose::TypeConstraints qw/duck_type/;
use YAML;
use Clone           qw/clone/;

=begin TODO

  - triggers
  - install as DBI callback for replacement of DBI::col_info())
  - multi-owners ... or build separate class to compare/merge owners.


=cut

our $VERSION = 0.1;

#======================================================================
# ATTRIBUTES AND BUILDER METHODS
#======================================================================


has 'dbh'              => (is => 'ro', isa => 'DBI::db', required   => 1);
has 'owner'            => (is => 'ro', isa => 'Str',     lazy_build => 1);
has 'log'              => (is => 'ro',
                           isa     => duck_type([qw/debug info/]),
                           handles => [qw/debug info/],
                           );

has 'want_constraints' => (is => 'ro', isa => 'Bool',    default => 1);
has 'want_index'       => (is => 'ro', isa => 'Bool',    default => 1);
has 'want_comments'    => (is => 'ro', isa => 'Bool',    default => 1);

sub _build_owner {
  my $self = shift;
  return $self->dbh->{Username};
}


#======================================================================
# PUBLIC METHODS
#======================================================================


sub tables { # not modelled as an attribute because of lazy loading from Oracle
             # and auto-vivification in $self->_add_to_col()
  my $self = shift;

  if (!$self->{tables}) {
    $self->_load_tab_cols;
    $self->_load_constraints if $self->want_constraints;
    $self->_load_index_info  if $self->want_index;
    $self->_load_comments    if $self->want_comments;

    $self->log->info("DICTIONARY LOADED") if $self->log;
  }

  return $self->{tables};
}


sub warnings {
  my $self = shift;
  return $self->{warnings} // {};
}


#======================================================================
# INTERNAL METHODS FOR LOADING ORACLE DATA DICITIONARY
#======================================================================


sub _load_tab_cols {
  my $self = shift;
  my $data = {};

  $self->log->info("loading ALL_TAB_COLS") if $self->log;

  my @fields = qw/COLUMN_NAME DATA_TYPE CHAR_LENGTH NULLABLE DATA_DEFAULT
                  DATA_PRECISION DATA_SCALE VIRTUAL_COLUMN/;

  # Note : using ALL_TAB_COLS instead of ALL_TAB_COLUMNS because
  # VIRTUAL_COLUMN is only in ALL_TAB_COLS
  local $" = ", "; # LIST_SEPARATOR
  my $cols = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT TABLE_NAME, @fields
      FROM ALL_TAB_COLS
     WHERE OWNER=? AND HIDDEN_COLUMN='NO'

  foreach my $col (@$cols) {
    $self->_add_to_col($col, $_ => $col->{$_}) foreach @fields;
  }
}

sub _load_constraints {
  my $self = shift;
  my $data = {};

  $self->log->info("loading ALL_CONSTRAINTS") if $self->log;


  # Info for primary keys and foreign keys requires several joins between
  # ALL_CONSTRAINTS and ALL_CONS_COLUMNS. For faster results, we first download
  # all data and then perform the joins in memory.

  # hash of  ALL_CONSTRAINTS by CONSTRAINT_NAME
  my $constraint = $self->dbh->selectall_hashref(<<"", 'CONSTRAINT_NAME', {}, $self->owner);
    SELECT  CONSTRAINT_NAME, CONSTRAINT_TYPE, R_CONSTRAINT_NAME
      FROM  ALL_CONSTRAINTS 
      WHERE OWNER = ?
        AND CONSTRAINT_TYPE IN ('R', 'P', 'U')
        AND STATUS='ENABLED'

  $self->log->info("loading ALL_CONS_COLUMNS") if $self->log;

  # list of  ALL_CONS_COLUMNS
  my $cons_columns = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT  CONSTRAINT_NAME, TABLE_NAME, COLUMN_NAME, POSITION
      FROM  ALL_CONS_COLUMNS
      WHERE OWNER = ?

  foreach my $row (@$cons_columns) {
    if (my $pos = $row->{POSITION}) {
      $constraint->{$row->{CONSTRAINT_NAME}}{cols}[$pos-1] = $row;
    }
    else {
      $self->_warn(sprintf "UNEXPECTED ALL_CONS_COL %s (%s.%s)",
                   @{$row}{qw/CONSTRAINT_NAME TABLE_NAME COLUMN_NAME/});
    }
  }

  # build internal data structure
 CONSTRAINT:
  foreach my $c (values %$constraint) {
    if ($c->{CONSTRAINT_TYPE} eq 'P') { # primary key
      $self->_add_to_col($_, is_prim_key => 1) foreach @{$c->{cols}};
    }
    if ($c->{CONSTRAINT_TYPE} eq 'U') { # unique
      $self->_add_to_col($_, is_unique => 1) foreach @{$c->{cols}};
    }
    elsif ($c->{CONSTRAINT_TYPE} eq 'R') { # foreign key
      my $i = 0;
      foreach my $col (@{$c->{cols}}) {
        my $r_name            = $c->{R_CONSTRAINT_NAME};
        my $remote_constraint = $constraint->{$r_name}
          or $self->_warn("NO CONSTRAINT $r_name") and next CONSTRAINT;
        my $r_col = $remote_constraint->{cols}[$i++]
          or $self->_warn("NO COLS IN  $r_name") and next CONSTRAINT;
        $self->_add_to_col($col,   references     => $r_col);
        $self->_add_to_col($r_col, is_referred_by => $col, 'PUSH');
      }
    }
  }
}



sub _load_index_info {
  my $self = shift;

  $self->log->info("loading ALL_INDEXES/ALL_IND_COLUMNS") if $self->log;

  my $ix_cols = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT INDEX_TYPE, ALL_IND_COLUMNS.TABLE_NAME, COLUMN_NAME
      FROM ALL_IND_COLUMNS JOIN ALL_INDEXES USING (INDEX_NAME)
     WHERE ALL_IND_COLUMNS.TABLE_OWNER=?

  $self->_add_to_col($_, INDEX_TYPE => $_->{INDEX_TYPE}) foreach @$ix_cols;
}



sub _load_comments {
  my $self = shift;

  $self->log->info("loading COMMENTS") if $self->log;

  # table comments
  my $tab_comments = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT TABLE_NAME, COMMENTS
      FROM ALL_TAB_COMMENTS
     WHERE OWNER=?

  $self->_add_to_table($_, COMMENTS => $_->{COMMENTS}) foreach @$tab_comments;

  # column comments
  my $col_comments = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT TABLE_NAME, COLUMN_NAME, COMMENTS
      FROM ALL_COL_COMMENTS
     WHERE OWNER=?

  $self->_add_to_col($_, COMMENTS => $_->{COMMENTS}) foreach @$col_comments;
}


#======================================================================
# UTILITY METHODS
#======================================================================



sub _add_to_table {
  my ($self, $hash, $key, $val) = @_;
  my $tname = $hash->{TABLE_NAME};

  # if this is a new table, store its name at the table level
  $self->{tables}{$tname}{TABLE_NAME} = $tname if !exists $self->{tables}{$tname};

  $self->{tables}{$tname}{$key} = $val         if $key;

  return $self->{tables}{$tname};
}



sub _add_to_col {
  my ($self, $hash, $key, $val, $push) = @_;

  # if $val is a subtree, clone it
  $val = clone $val if ref $val;

  my $table = $self->_add_to_table($hash);
  my $cname = $hash->{COLUMN_NAME};

  # insert into datatree
  if ($push) {
    push @{$table->{col}{$cname}{$key}}, $val;
  }
  else {
    $table->{col}{$cname}{$key} = $val;
  }
}



sub _warn {
  my ($self, $warning) = @_;

  $self->{warnings}{$warning} +=1;
}





1;

__END__

=head1 NAME

Oracle::DataDictionary

=head1 SYNOPSIS

  use DBI;
  use Oracle::DataDictionary;
  
  my $dbh = DBI->connect("dbi:Oracle:...", ...);
  my dict = Oracle::DataDictionary->new(dbh => $dbh);

  my tables = $dict->tables;
  foreach my $table (values %$tables) {
    print $table->{TABLE_NAME}, "\n";
    foreach my $col (values %{$table->{col}}) {
       print "  $col->{COLUMN_NAME} : $col->{DATA_TYPE}($col->{CHAR_LENGTH})\n";
    }
  }


=head1 DESCRIPTION

This module reads Oracle data dictionary tables (ALL_TAB_COLS, ALL_CONSTRAINTS, etc.)
and assembles the collected information into a global datatree, suitable for export
for example in JSON, YAML or XML format.

The companion module L<Oracle::DataDictionary::ToMsWord> uses the datatree to
produce complete documentation of an Oracle schema as a Microsoft Word document.

=head1 METHODS

=head2 new

  my dict = Oracle::DataDictionary->new(dbh => $dbh, ...);

Creates an instance of a dictionary. Arguments to the C<new()> method are:

=over

=item dbh

mandatory. L<DBI> database handle to the Oracle instance.

=item owner.

optional. Name of the schema owner. Default is the connected user.

=item options to control the amount of downloaded data

Options below are all true by default. They can be set to false if the
information is not needed; this would spare some download time and some memory.

=over

=item want_constraints

download data about Oracle constraints (primary and foreign keys, unicity)

=item want_index

download data about the types of indexes associated to columns

=item want_comments

download comments associated to tables and columns

=back

=back



=head2 tables

Returns a hashref where keys are table names and values are nested hashrefs with :

=over

=item TABLE_NAME

=item COMMENTS

=col

=back

The C<col> entry is a hashref where keys are column names and values
are nested hashrefs with :

=over

=item COLUMN_NAME

=item DATA_TYPE

Oracle datatype

=item CHAR_LENGTH

For VARCHAR2 columns : number of characters

=item NULLABLE

'O' if the column accepts null, 'N' otherwise

=item DATA_DEFAULT

Expression to compute either a default value for the column,
or a computed value for a virtual column (see C<VIRTUAL_COLUMN> below).

=item DATA_PRECISION

Length in decimal digits (NUMBER) or binary digits (FLOAT)

=item DATA_SCALE

Digits to the right of the decimal point in a number

=item VIRTUAL_COLUMN

'YES' if the column is virtual (value dynamically computed from the
C<DATA_DEFAULT> expression); 'NO' otherwise

=item COMMENTS

=item INDEX_TYPE

If non-null, the column is indexed. See Oracle documentation for the different
types of indexes. Columns with a fulltext index (Oracle Text component)
are reported as 'DOMAIN' index type.


=item is_prim_key

true if this column is (or belongs to) a primary key

=item is_unique

true if this column has a uniqueness constraint

=item references

if this column is a foreign key, hashref to a copy of the column
that contains the referred primary key

=item is_referred_by

arrayref of copies of columns that contain foreign keys to the current column


=back


=head2 warnings

Returns a hashref of warnings encountered while downloading Oracle data dictionary
information. Keys or the hashref are textual warning messages; values of the hashref
are the number of occurrences for each message.

Possible messages are :

=over

=item UNEXPECTED ALL_CONS_COL

An entry was found in ALL_CONS_COLUMNS with no corresponding constraint
in ALL_CONSTRAINTS

=item NO CONSTRAINT <remote_constraint_name>

The "remote constraint" referenced in an ALL_CONS_COLUMNS entry was not found.

=item NO COLS IN  <remote_constraint_name>

The "remote constraint" referenced in an ALL_CONS_COLUMNS entry was found
but contains no column descriptions.

=back


=head1 AUTHOR

Laurent Dami, C<< <laurent dot dami at cpan dot org> >>, May 2020.



=head1 LICENSE AND COPYRIGHT

Copyright 2020 Laurent Dami.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

