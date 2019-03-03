using System.Linq;
using ICSharpCode.CodeConverter.Util;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using SyntaxFactory = Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using SyntaxKind = Microsoft.CodeAnalysis.CSharp.SyntaxKind;
using VBSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax;
using VBasic = Microsoft.CodeAnalysis.VisualBasic;

namespace ICSharpCode.CodeConverter.CSharp
{

    /// <summary>
    /// The equals and not equals operators in Visual Basic call ConditionalCompareObjectEqual.
    /// This method allows a sort of best effort comparison of different types.
    /// There are therefore some surprising results such as "" = Nothing being true.
    /// Here we try to coerce the arguments for the CSharp equals method to get as close to the runtime behaviour as possible without inlining hundreds of lines of code.
    /// </summary>
    internal class VisualBasicStringComparison
    {
        public static (ExpressionSyntax lhs, ExpressionSyntax rhs) CoerceEqualityArgs(SemanticModel semanticModel,
            VBSyntax.BinaryExpressionSyntax node, ExpressionSyntax lhs, ExpressionSyntax rhs)
        {
            var leftType = semanticModel.GetTypeInfo(node.Left).Type;
            var rightType = semanticModel.GetTypeInfo(node.Right).Type;

            if (new[] { leftType, rightType }.All(t =>
                t.SpecialType == SpecialType.System_String ||
                t.SpecialType == SpecialType.System_Object ||
                t.IsArrayOf(SpecialType.System_Char))) {
                lhs = VbCoerceToString(semanticModel, lhs, leftType);
                rhs = VbCoerceToString(semanticModel, rhs, rightType);
            }

            return (lhs, rhs);
        }

        private static ExpressionSyntax VbCoerceToString(SemanticModel semanticModel, ExpressionSyntax baseExpression,
            ITypeSymbol type)
        {
            bool isStringType = type.SpecialType == SpecialType.System_String;

            var constantValue = semanticModel.GetConstantValue(baseExpression);
            if (!constantValue.HasValue) {
                if (isStringType) {
                    baseExpression = SyntaxFactory.ParenthesizedExpression(Coalesce(baseExpression, EmptyStringExpression()));
                } else {
                    var baseExpressionAsCharArray = SyntaxFactory.BinaryExpression(
                        SyntaxKind.AsExpression,
                        SyntaxFactory.ParenthesizedExpression(baseExpression),
                        CharArrayType().WithRankSpecifiers(ArrayRankSpecifier(SyntaxFactory.OmittedArraySizeExpression())));
                    baseExpression = Coalesce(baseExpressionAsCharArray, EmptyCharArrayExpression());
                }
            } else if (constantValue.Value == null) {
                baseExpression = EmptyStringExpression();
            }

            return isStringType ? baseExpression : NewStringFromArg(baseExpression);
        }

        private static ExpressionSyntax Coalesce(ExpressionSyntax lhs, ExpressionSyntax emptyString)
        {
            return SyntaxFactory.BinaryExpression(SyntaxKind.CoalesceExpression, lhs, emptyString);
        }

        private static ArrayCreationExpressionSyntax EmptyCharArrayExpression()
        {
            var literalExpressionSyntax =
                SyntaxFactory.LiteralExpression(SyntaxKind.NumericLiteralExpression, SyntaxFactory.Literal(0));
            var arrayRankSpecifierSyntax = ArrayRankSpecifier(literalExpressionSyntax);
            var arrayTypeSyntax = CharArrayType().WithRankSpecifiers(arrayRankSpecifierSyntax);
            return SyntaxFactory.ArrayCreationExpression(arrayTypeSyntax);
        }

        private static SyntaxList<ArrayRankSpecifierSyntax> ArrayRankSpecifier(ExpressionSyntax expressionSyntax)
        {
            var literalExpressionSyntaxList = SyntaxFactory.SingletonSeparatedList(expressionSyntax);
            var arrayRankSpecifierSyntax = SyntaxFactory.ArrayRankSpecifier(literalExpressionSyntaxList);
            return SyntaxFactory.SingletonList(arrayRankSpecifierSyntax);
        }

        private static ArrayTypeSyntax CharArrayType()
        {
            return SyntaxFactory.ArrayType(SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.CharKeyword)));
        }

        private static LiteralExpressionSyntax EmptyStringExpression()
        {
            return SyntaxFactory.LiteralExpression(SyntaxKind.StringLiteralExpression, SyntaxFactory.Literal(""));
        }

        private static ObjectCreationExpressionSyntax NewStringFromArg(ExpressionSyntax lhs)
        {
            return SyntaxFactory.ObjectCreationExpression(
                SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.StringKeyword)),
                ExpressionSyntaxExtensions.CreateArgList(lhs), default(InitializerExpressionSyntax));
        }
    }
}