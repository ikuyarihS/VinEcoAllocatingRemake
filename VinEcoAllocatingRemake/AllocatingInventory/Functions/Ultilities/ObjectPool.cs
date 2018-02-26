using System;
using System.Collections.Concurrent;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    #region

    #endregion

    /// <summary>
    ///     The object pool.
    /// </summary>
    /// <typeparam name="T">
    ///     Type in pool.
    /// </typeparam>
    public class ObjectPool<T>
    {
        /// <summary>
        ///     The object generator.
        /// </summary>
        private readonly Func<T> _objectGenerator;

        /// <summary>
        ///     The objects.
        /// </summary>
        private readonly ConcurrentBag<T> _objects;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ObjectPool{T}" /> class.
        /// </summary>
        /// <param name="objectGenerator">
        ///     The object generator.
        /// </param>
        public ObjectPool(Func<T> objectGenerator)
        {
            _objects         = new ConcurrentBag<T>();
            _objectGenerator = objectGenerator ?? throw new ArgumentNullException(nameof(objectGenerator));
        }

        /// <summary>
        ///     The get object.
        /// </summary>
        /// <returns>
        ///     The <see cref="T" />.
        /// </returns>
        public T GetObject()
        {
            return _objects.TryTake(out T item)
                       ? item
                       : _objectGenerator();
        }

        /// <summary>
        ///     The put object.
        /// </summary>
        /// <param name="item">
        ///     The item.
        /// </param>
        public void PutObject(T item)
        {
            _objects.Add(item);
        }
    }
}