using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq.Expressions;

namespace Npoi.Mapper
{
    /// <summary>
    /// This class is not used currently.
    /// Used for dynamic object pattern, it's only for dynamic data type.
    /// </summary>
    internal class RowDataDynamic : IDynamicMetaObjectProvider, IDictionary<string, object>
    {
        private readonly IDictionary<string, object> _dictionary;

        public RowDataDynamic() : this(0) { }

        public RowDataDynamic(int capacity)
        {
            _dictionary = new Dictionary<string, object>(capacity);
        }
        /// <inheritdoc />
        public DynamicMetaObject GetMetaObject(Expression parameter)
        {
            return new RowDataMetaObject(parameter, BindingRestrictions.Empty, this);
        }

        public void SetValue(string name, object value)
        {
            _dictionary[name] = value;
        }

        #region Implementation of IDictionary<string,object>

        /// <inheritdoc />
        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc />
        public void Add(KeyValuePair<string, object> item)
        {
            _dictionary.Add(item);
        }

        /// <inheritdoc />
        public void Clear()
        {
            _dictionary.Clear();
        }

        /// <inheritdoc />
        public bool Contains(KeyValuePair<string, object> item)
        {
            return _dictionary.Contains(item);
        }

        /// <inheritdoc />
        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex)
        {
            _dictionary.CopyTo(array, arrayIndex);
        }

        /// <inheritdoc />
        public bool Remove(KeyValuePair<string, object> item)
        {
            return _dictionary.Remove(item);
        }

        /// <inheritdoc />
        public int Count => _dictionary.Count;

        /// <inheritdoc />
        public bool IsReadOnly => _dictionary.IsReadOnly;

        /// <inheritdoc />
        public bool ContainsKey(string key)
        {
            return _dictionary.ContainsKey(key);
        }

        /// <inheritdoc />
        public void Add(string key, object value)
        {
            _dictionary.Add(key, value);
        }

        /// <inheritdoc />
        public bool Remove(string key)
        {
            return _dictionary.Remove(key);
        }

        /// <inheritdoc />
        public bool TryGetValue(string key, out object value)
        {
            return _dictionary.TryGetValue(key, out value);
        }

        /// <inheritdoc />
        public object this[string key]
        {
            get { return _dictionary[key]; }
            set { _dictionary[key] = value; }
        }

        /// <inheritdoc />
        public ICollection<string> Keys => _dictionary.Keys;

        /// <inheritdoc />
        public ICollection<object> Values => _dictionary.Values;

        #endregion
    }
}
