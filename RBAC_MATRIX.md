# RBAC MATRIX

| Capability | Admin | Operator | Viewer |
|---|---|---|---|
| `setupSystem()` | Yes | No | No |
| Mo `START_HERE` | Yes | Yes | Yes |
| Mo `TrainingEntry` | Yes | Yes | No |
| Mo `ImportCenter` | Yes | Yes | No |
| Upload vao staging | Yes | Yes | No |
| Chay QA | Yes | Yes | No |
| Publish staging | Yes | Yes | No |
| Lam moi analytics | Yes | Yes | No |
| Xem dashboards | Yes | Yes | Yes |
| Xem analytics | Yes | Yes | Yes |
| Xem snapshots | Yes | Yes | Yes |
| Sua `MASTER_*` | Yes | Yes | No |
| Sua `CONFIG_USERS` | Yes | No | No |
| Backup | Yes | No | No |
| Restore | Yes | No | No |
| Archive nam | Yes | No | No |
| Xem log | Yes | No | No |

## Notes
- Neu `CONFIG_USERS` rong, nguoi dau tien chay `setupSystem()` duoc seed thanh `Admin`.
- Nguoi khong co role hop le mac dinh la `Viewer`.
