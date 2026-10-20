using BrimOutlook.Outlook;

namespace BrimOutlook;

/// <summary>
/// Kurze Schlüssel für Outlook-Elemente in Aktions-IDs („open:i7“). EntryIDs sind über 100 Zeichen lang;
/// außerdem bleibt so nichts aus Outlook in Brims Protokoll stehen.
/// Aktionen kommen auf anderen Threads an als das Zeichnen, daher mit Sperre.
/// </summary>
public sealed class ActionKeys
{
    private const int Limit = 500;

    private readonly Lock _gate = new();
    private readonly Dictionary<ItemRef, string> _keys = [];
    private readonly Dictionary<string, ItemRef> _items = [];
    private int _next;

    /// <summary>Schlüssel für ein Element; gleiche Elemente bekommen immer denselben.</summary>
    public string For(ItemRef item)
    {
        using var scope = _gate.EnterScope();
        if (_keys.TryGetValue(item, out var key))
        {
            return key;
        }
        key = "i" + (++_next).ToString(System.Globalization.CultureInfo.InvariantCulture);
        _keys[item] = key;
        _items[key] = item;
        return key;
    }

    public ItemRef? Resolve(string key)
    {
        using var scope = _gate.EnterScope();
        return _items.GetValueOrDefault(key);
    }

    /// <summary>Vergisst alles außer <paramref name="current"/>, sobald sich zu viel angesammelt hat.</summary>
    public void Trim(IEnumerable<ItemRef> current)
    {
        using var scope = _gate.EnterScope();
        if (_keys.Count <= Limit)
        {
            return;
        }
        var keep = current.Where(_keys.ContainsKey).ToDictionary(item => item, item => _keys[item]);
        _keys.Clear();
        _items.Clear();
        foreach (var (item, key) in keep)
        {
            _keys[item] = key;
            _items[key] = item;
        }
    }
}
