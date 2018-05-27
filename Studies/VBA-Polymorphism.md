#Polymorphism in VBA

From my understanding, VBA programmers are to use **Implements** to implement Polymorphism:

Similar to [An example](http://realanalysiszone.com/?p=281)
```VBA
'IMonster interface
Public Property Get Name() As String
     
End Property
 
Public Property Let Name(Value As String)
     
End Property
 
Public Property Get HP() As Integer
 
End Property
 
Public Property Let JP(Value As Integer)
     
End Property
```
And then we can **Implements** the interface:
```VBA
'Goblin
Public Property Get IMonster_Name() As String
  IMonster_Name = m_Name
End Property

Public Property Let IMonster_Name(Value As String) As String
  m_Name = Value
End Property
'Rest of the properties
```

And another monster:
```VBA
'Orc
Public Property Get IMonster_Name() As String
  IMonster_Name = m_Name
End Property

Public Property Let IMonster_Name(Value As String) As String
  m_Name = Value
End Property
'Rest of the properties
```

Then we can **Switch** to generate monsters:
```VBA
Sub CreateMonster(ByVal monster As String, ByVal name As String, ByVal hp As Integer)
  Dim oTemplate As IMonster
  Select Case monster
    Case "Goblin"
      Set oTemplate = New Goblin
      oTemplate.Name = name
      oTemplate.HP = hp
    Case "Orc"
      Set oTemplate = New Orc
      oTemplate.Name = name
      oTemplate.HP = hp
    Case Else
      Set oTemplate = Nothing
  End Select
End Sub
```

Any programmer with some experience will immediately recognize that the potential problem of spamming **Case**

I happened to stump upon this piece of code:
(Better framework)[https://codereview.stackexchange.com/questions/64109/extensible-logging]

The author cleverly used a **Dictionary** to serve as a storage for all Loggers, the user may create his own Logger, by **Implements**
the ILogger Interface, and registering it by calling the **Register** function of the LogManager.
```VBA
LogManager.Register DebugLogger.Create("MyLogger", DebugLevel)
```
And call the **Log** function in LogManager to format and output the log info:
```
LogManager.Log DebugLevel, "we're done here.", "TestLogger"
```
You can see that this is very neat and is closer to C++'s Polymorphism.
