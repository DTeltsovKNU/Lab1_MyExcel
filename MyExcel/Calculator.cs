using Antlr4.Runtime;
using System;
namespace MyExcel
{
public static class Calculator
    {
        public static double Evaluate(string expression)
        {
            var lexer = new MyExcelLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());
            var tokens = new CommonTokenStream(lexer);
            var parser = new MyExcelParser(tokens);
            var tree = parser.compileUnit();
            var visitor = new MyExcelVisitor();
            return visitor.Visit(tree);
        }
    }
}
