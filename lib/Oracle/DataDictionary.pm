package Oracle::DataDictionary;
use Moose;
use YAML;
use Clone           qw/clone/;

=begin TODO

  - triggers

  - Ora::DD::ToMsWord
       - several owners
       - 

  - as_DBI_col_info


   - THINK api "tables" // {table}

=cut


has 'dbh'              => (is => 'ro', isa => 'DBI::db', required   => 1);
has 'owner'            => (is => 'ro', isa => 'Str',     lazy_build => 1);

has 'want_constraints' => (is => 'ro', isa => 'Bool',    default => 1);
has 'want_index'       => (is => 'ro', isa => 'Bool',    default => 1);
has 'want_comments'    => (is => 'ro', isa => 'Bool',    default => 0);

has 'tables'           => (is => 'ro', isa => 'HashRef', lazy_build => 1);



sub _build_owner {
  my $self = shift;
  return $self->dbh->{Username};
}

sub _build_tables {
  my $self = shift;

  $self->_load_col_info;
  $self->_load_constraints if $self->want_constraints;
  $self->_load_index_info  if $self->want_index;
  $self->_load_comments    if $self->want_comments;

  warn scalar(localtime), " DONE\n";

  return $self->{table};
}



sub _load_col_info {
  my $self = shift;
  my $data = {};

  warn scalar(localtime), " ALL_TAB_COLS\n";

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

  warn scalar(localtime), " ALL_CONSTRAINTS\n";


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

  warn scalar(localtime), " ALL_CONS_COLUMNS\n";

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
      $self->add_warning(sprintf "UNEXPECTED ALL_CONS_COL %s (%s.%s)",
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
          or $self->add_warning("NO CONSTRAINT $r_name") and next CONSTRAINT;
        my $r_col = $remote_constraint->{cols}[$i++]
          or $self->add_warning("NO COLS IN  $r_name") and next CONSTRAINT;
        $self->_add_to_col($col,   references     => $r_col);
        $self->_push_to_col($r_col, is_referred_by => $col  );
      }
    }
  }
}



sub _load_index_info {
  my $self = shift;

  warn scalar(localtime), " INDEX\n";

  my $ix_cols = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT INDEX_TYPE, ALL_IND_COLUMNS.TABLE_NAME, COLUMN_NAME
      FROM ALL_IND_COLUMNS JOIN ALL_INDEXES USING (INDEX_NAME)
     WHERE ALL_IND_COLUMNS.TABLE_OWNER=?

  $self->_add_to_col($_, index_type => $_->{INDEX_TYPE}) foreach @$ix_cols;
}



sub _load_comments {
  my $self = shift;

  warn scalar(localtime), " COMMENTS\n";

  my $comments = $self->dbh->selectall_arrayref(<<"", {Slice => {}}, $self->owner);
    SELECT TABLE_NAME, COLUMN_NAME, COMMENTS
      FROM ALL_COL_COMMENTS
     WHERE OWNER=?

  $self->_add_to_col($_, comments => $_->{COMMENTS}) foreach @$comments;
}


sub _add_to_col {
  my ($self, $col, $key, $val) = @_;
  my $tname = $col->{TABLE_NAME};
  my $cname = $col->{COLUMN_NAME};

  $self->{table}{$tname}{TABLE_NAME} = $tname if !exists $self->{table}{$tname};
  $self->{table}{$tname}{col}{$cname}{$key} = ref $val ? clone $val : $val;
}


sub _push_to_col {
  my ($self, $col, $key, $val) = @_;
  my $tname = $col->{TABLE_NAME};
  my $cname = $col->{COLUMN_NAME};

  push @{$self->{table}{$tname}{col}{$cname}{$key}}, ref $val ? clone $val : $val;
}



sub add_warning {
  my ($self, $warning) = @_;

  $self->{warnings}{$warning} +=1;
}


sub warnings {
  my $self = shift;
  $self->{warnings} // {};
}



1;

__END__











  my $comments = $self->{dbh}->selectall_hashref(<<"", 'COLUMN_NAME');
    SELECT COLUMN_NAME, COMMENTS
      FROM ALL_COL_COMMENTS
     WHERE OWNER='$self->{owner}'
       AND TABLE_NAME='$table'
     ORDER BY COLUMN_NAME

  my ($all_triggers, $autoincrements) = $self->triggers($table);

  my @col_descrs;
  my @all_fk_indices;
  my @all_comments;
  foreach my $col (@cols) {
    my $col_name = $col->{COLUMN_NAME};

    my ($col_descr, $fk_indices, $triggers) 
      = $self->col_info($table, $col, $prim_key, $indices, $autoincrements);

    # insertion d'un séparateur pour les champs techniques
    $col_descr = "\n  ----------$col_descr" if $col->{COLUMN_NAME} eq 'OPER_CREAT';

    push @col_descrs, $col_descr;
    push @all_fk_indices, @$fk_indices;
    push @$all_triggers, @$triggers;

    if ($col_name !~ /CREAT|MODIF/ and my $comment = $comments->{$col_name}{COMMENTS}) {
      my $spaces = ' ' x (length($col_name) + 3);
      $comment =~ s/\n/\n$spaces/g;
      push @all_comments, "$col_name : $comment";
    }
  }



  # TODO :
  # - fulltext
  # - comments
}


sub col_info {
  my ($self, $table, $col, $prim_key, $indices, $autoincrements) = @_;

  my $col_name = $col->{COLUMN_NAME};
  my @fk_indices;

  # description de la colonne, avec index unique le cas échéant
  my $type = $col->{DATA_TYPE};
  $type    = 'INTEGER' if $type eq 'NUMBER' and $col_name !~ /^MNT_/;

  # taille pour les champs numériques
  if ($type =~ /^(INTEGER|NUMBER)$/) {
    my $precision;
    $precision = $col->{DATA_PRECISION}   if $col->{DATA_PRECISION};
    $precision .= ",$col->{DATA_SCALE}"   if $col->{DATA_SCALE};
    $type .= "($precision)" if $precision;
  }

  # taille pour les champs alphabétiques
  $type   .= "($col->{CHAR_LENGTH})" if $col->{CHAR_LENGTH};

  my $index = "";
  if (my $ix = $indices->{$col_name}) {
    if ($ix->{INDEX_TYPE} eq 'NORMAL') {
      if ($ix->{UNIQUENESS} eq 'UNIQUE') {
        if ($col_name eq $prim_key) {
          $index = "PRIMARY KEY";
          $index .= ' AUTOINCREMENT' if any { $col_name eq $_ } @$autoincrements;
        }
        else {
          $index = "UNIQUE";
        }
      } 
      else {
        push @fk_indices,
          sprintf "CREATE INDEX %-30s ON $table($col_name);",
                  $ix->{INDEX_NAME};
      }
    }
    elsif ($ix->{INDEX_TYPE} eq 'DOMAIN') {
      $index = "/* fulltext */";
    }

  }


}


# NOTE : les "DEFAULT" de Oracle sont soit de vraies valeurs par défaut,
# soit les expressions associées aux colonnes virtuelles


sub parse_default {
  my ($self, $table, $col) = @_;
  my $cname = $col->{COLUMN_NAME};
  my $default = "";
  my @trigger;

  if (any { $cname eq $_ } qw/TS_CREAT TS_MODIF/) {
    $default = "DEFAULT CURRENT_TIMESTAMP";
  }
  elsif (any { $cname eq $_ } qw/OPER_CREAT OPER_MODIF/) {
    $default = "DEFAULT 'SQLITE'";
  }
  elsif ($cname eq 'D_MODIF') {
    $default = "DEFAULT CURRENT_DATE";
  }
  elsif ($cname eq 'T_MODIF') {
    $default = "DEFAULT CURRENT_TIME";
  }
  elsif (    $cname =~ /^[DT]_(CONC_|DEST_)?DATE[12]/
          || $cname =~ /^[DT]_(DEB|FIN)_PLANIF/
          || $cname =~ /^[DT]_(EDIT|EXPED|FIN)_ENVOI/
          || $cname =~ /^[DT]_REMISE_LOT/
          || $cname =~ /^[DT]_SUIVI_ENVOI/
          || (any { $cname eq $_ } qw/N_EXT_DECIS ATTR_EXT N_EXT_JUR_ATTR N_EXT_MESURE/)
        ) {
    # ignore, handled by hardcoded triggers
  }
  else {
    my $data_def = $col->{DATA_DEFAULT} || '';
    if ($data_def =~ /"F_(.*?)"\("(.*?)"\)/) { # fonction PL/SQL de traduction de code-lists
      push @trigger, $self->build_CL_trigger($1, $table, $2, $cname);
    }
    elsif ($data_def =~ /NOM_/ && $data_def =~ /PRE_/) { # concaténation des noms/prénoms
      push @trigger, $self->build_NOM_PRE_trigger($table, $cname);
    }
    elsif ($data_def =~ /\(/) {             # autre expression PL/SQL, mise en commentaire
      $default = "\n    /* DEFAULT $data_def */";
    }
    elsif ($data_def =~ /./) {              # constante, reconduite telle quelle
      $default = "DEFAULT $data_def";
    }
  }

  return ($default, @trigger);
}

