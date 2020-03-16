{$IFDEF VER150}
  {$DEFINE VTK_D7}
{$ELSE}
  {$IFDEF VER140}
    {$DEFINE VTK_D6}
    {$IFDEF BCB}
      {$DEFINE VTK_CB6}
    {$ENDIF}
  {$ELSE}
    {$IFDEF VER130}
      {$DEFINE VTK_D5}
      {$IFDEF BCB}
        {$DEFINE VTK_CB5}
      {$ENDIF}
    {$ELSE}
      {$IFDEF VER120}
        {$DEFINE VTK_D4}
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}
{$ENDIF}

//{$DEFINE DEBUG}
