/** Constrói um UPDATE SQL a partir de cláusulas já validadas pela aplicação. */
class SqlStatementBuilder {
  constructor(tableName, escapeIdentifier, formatValue) {
    this.tableName = tableName;
    this.escapeIdentifier = escapeIdentifier;
    this.formatValue = formatValue;
    this.setClauses = [];
    this.whereClauses = [];
  }

  set(field, value) {
    const identifier = this.escapeIdentifier(field);
    const clause = `${identifier} = ${this.formatValue(value)}`;
    const index = this.setClauses.findIndex(item => item.startsWith(`${identifier} =`));
    if (index === -1) this.setClauses.push(clause);
    else this.setClauses[index] = clause;
  }

  where(field, value) {
    this.whereClauses.push(`${this.escapeIdentifier(field)} = ${this.formatValue(value)}`);
  }

  get canBuild() {
    return this.setClauses.length > 0 && this.whereClauses.length > 0;
  }

  build() {
    if (!this.canBuild) return null;
    return `UPDATE ${this.escapeIdentifier(this.tableName)} SET ${this.setClauses.join(', ')} WHERE ${this.whereClauses.join(' AND ')};`;
  }
}

window.SqlStatementBuilder = SqlStatementBuilder;
