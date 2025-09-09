if [km final] >= [km inicial] then
    List.Numbers([km inicial], [km final] - [km inicial] + 1, 1)
else
    List.Numbers([km inicial], [km inicial] - [km final] + 1, -1)
