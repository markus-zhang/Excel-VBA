**Factory Pattern in VBA**
The key is to:
- Make sure that clients are not able to modify the properties of the class/type object. This can be achieved by using `Friend` before `Let`.
- Code a factory class in the same project. Manually put VB_PredeclaredId to `True`.
- Use a `CreateItem` function in the factory class to return an item class/type object with properties set by parameters.

**Example**
```VBA
Private Type tGoblin
  iHP as Integer
  iDamage as Integer
  strName as String
End Type

Private this as tGoblin

'Omit other properties
Friend Property Let iHP(ByVal hp As Integer)
    this.iHP = hp
End Property
'Other Friend Let properties

'In the same VBA project, a factory class may do this
 Public Function CreateGoblin(ByVal hp As Integer, ByVal damage As Integer, ByVal name As String) As tGoblin
     Dim result As New tGoblin
     result.iHP = hp
     result.iDamage = damage
     result.strName = name
     Set CreateGoblin = result
 End Function
 
 'Because the factory class has VB_PredeclaredId as True, clients outside of the VBA project may do this
 Dim mGoblin As tGoblin
 Set mGoblin = MonsterFactory.CreateGoblin(5, 2, "Goblin Archer")
```

Now anyone who has implemented Factory pattern in another language, say C++, will recognize the limitation of the code above.
If a game programmer wants to `Create` monsters in the factory, he has to implement functions for all monsters. He will need
`CreateGoblin()`, `CreateLich()`, etc.

What we want to achieve is true Factory pattern, such that we can do this:
```
Dim mGoblin As tMonster
Set mGoblin = MonsterFactory.CreateMonster(5, 2, "Goblin Archer", "Goblin")
```
In one sentence, we want Polymorphism.
