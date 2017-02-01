using System.Collections.Generic;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;

namespace Npoi.Mapper
{
    /// <summary>
    /// This class is not used currently.
    /// Used for dynamic object pattern, it's only for dynamic data type.
    /// </summary>
    internal class RowDataMetaObject : DynamicMetaObject
    {
        private static readonly MethodInfo GetValueMethod = typeof(IDictionary<string, object>).GetProperty("Item").GetGetMethod();
        private static readonly MethodInfo SetValueMethod = typeof(RowDataDynamic).GetMethod("SetValue", new[] { typeof(string), typeof(object) });

        /// <inheritdoc />
        public RowDataMetaObject(Expression expression, BindingRestrictions restrictions) : base(expression, restrictions)
        {
        }

        /// <inheritdoc />
        public RowDataMetaObject(Expression expression, BindingRestrictions restrictions, object value) : base(expression, restrictions, value)
        {
        }

        private DynamicMetaObject CallMethod(MethodInfo method, Expression[] parameters)
        {
            return new DynamicMetaObject(
                Expression.Call(Expression.Convert(Expression, LimitType), method, parameters),
                BindingRestrictions.GetTypeRestriction(Expression, LimitType)
                );
        }

        public override DynamicMetaObject BindGetMember(GetMemberBinder binder)
        {
            return CallMethod(GetValueMethod, new Expression[] { Expression.Constant(binder.Name) });
        }

        // Needed for Visual basic dynamic support
        public override DynamicMetaObject BindInvokeMember(InvokeMemberBinder binder, DynamicMetaObject[] args)
        {
            return CallMethod(GetValueMethod, new Expression[] { Expression.Constant(binder.Name) });
        }

        public override DynamicMetaObject BindSetMember(SetMemberBinder binder, DynamicMetaObject value)
        {
            return CallMethod(SetValueMethod, new[] { Expression.Constant(binder.Name), value.Expression });
        }
    }
}
