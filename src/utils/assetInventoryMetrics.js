const normalizeStatus = (value) => String(value || '')
  .replace(/[\u200B-\u200D\uFEFF]/g, '')
  .trim();

export function buildVacantStockBreakdown(assets = [], getCategory = (itemType) => itemType) {
  return Array.from(assets
    .filter((asset) => normalizeStatus(asset.status) === 'ว่าง')
    .reduce((groups, asset) => {
      const label = getCategory(asset.itemType);
      groups.set(label, (groups.get(label) || 0) + 1);
      return groups;
    }, new Map()), ([label, count]) => ({
      label,
      displayLabel: label === 'Monitor' ? 'Monitor ว่าง' : label,
      count
    }))
    .sort((a, b) => {
      if (a.label === 'Monitor') return -1;
      if (b.label === 'Monitor') return 1;
      return b.count - a.count || a.label.localeCompare(b.label, 'th');
    });
}

