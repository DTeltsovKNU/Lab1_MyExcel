grammar MyExcel;


/*
 *Parser Rules
 */

compileUnit : expression EOF;

expression:
		'-' expression  #UnminExpr
		|LPAREN expression RPAREN #ParenthesizedExpr
		|expression EXPONENT expression #ExponentialExpr
	    |expression operatorToken=(MULTIPLY|DIVIDE) expression #MultiplicativeExpr
		|expression operatorToken=(MOD | DIV) expression #ModDivExpression
		|expression operatorToken=(ADD|SUBTRACT) expression #AdditiveExpr
		|NUMBER #NumberExpr
		|IDENTIFIER #identifierExpr
		|operatorToken=( MMIN | MMAX) LPAREN paramlist=arglist RPAREN #MinExpr
		;
arglist: expression (';' expression)+;

/*
 *Lexer Rules
 */

NUMBER : INT ('.'INT)?;
IDENTIFIER : [A-Z1-90-9]+;

INT : ('0'..'9')+;

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
MOD : 'mod';
DIV : 'div';
SUBTRACT : '-';
ADD : '+';
MMIN : 'mmin';
MMAX : 'mmax';
LPAREN : '(';
RPAREN : ')';

WS:[\t\r\n] -> channel(HIDDEN);