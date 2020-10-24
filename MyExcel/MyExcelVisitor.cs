using System.Diagnostics;
using System;
using System.Linq;
namespace MyExcel
{
    class MyExcelVisitor : MyExcelBaseVisitor<double>
    {
        public override double VisitCompileUnit(MyExcelParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitNumberExpr(MyExcelParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);
            return result;
        }


        public override double VisitIdentifierExpr(MyExcelParser.IdentifierExprContext context)
        {

            var result = context.GetText();
            double value = Convert.ToDouble(Form1.dict1[result].Value);
            //видобути значення змінної з таблиці
            if (Form1.dict1[result].Value != "")
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }


        public override double VisitParenthesizedExpr(MyExcelParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitUnminExpr(MyExcelParser.UnminExprContext context)
        {
            var number = WalkLeft(context);
            Debug.WriteLine("-{0}", number);
            return -number;
        }

        public override double VisitExponentialExpr(MyExcelParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }


        public override double VisitAdditiveExpr(MyExcelParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == MyExcelLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMinExpr(MyExcelParser.MinExprContext context)
        {

            if (context.operatorToken.Type == MyExcelLexer.MMIN)
            {
                double min = Double.PositiveInfinity;

                foreach (var child in context.paramlist.children.OfType<MyExcelParser.ExpressionContext>())
                {
                    double childValue = this.Visit(child);

                    if (childValue < min) min = childValue;
                }

                return min;
            }

            else
            {
                double max = Double.NegativeInfinity;

                foreach (var child in context.paramlist.children.OfType<MyExcelParser.ExpressionContext>())
                {
                    double childValue = this.Visit(child);

                    if (childValue > max) max = childValue;
                }

                return max;
            }
        }

        public override double VisitModDivExpression(MyExcelParser.ModDivExpressionContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (right == 0) throw new DivideByZeroException();
            else
            {
                if (context.operatorToken.Type == MyExcelLexer.MOD)
                {
                    Debug.WriteLine("{0} mod {1}", left, right);
                    return left % right;
                }

                else
                {
                    Debug.WriteLine("{0} div {1}", left, right);
                    return (int)(left / right);
                }
            }
        }

        public override double VisitMultiplicativeExpr(MyExcelParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == MyExcelLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                if (right == 0) throw new DivideByZeroException();
                else
                {
                    Debug.WriteLine("{0} / {1}", left, right);
                    return left / right;
                }
            }
        }


        private double WalkLeft(MyExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<MyExcelParser.ExpressionContext>(0));
        }


        private double WalkRight(MyExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<MyExcelParser.ExpressionContext>(1));
        }
    }
}
