DECLARE @var_OF VARCHAR(10)
SELECT @var_OF = Current_OF FROM Main_Table WHERE Id_Piece = (SELECT MAX(Id_Piece) from Main_Table)
SELECT Kogame_Data.Hauteur 
FROM  Kogame_Data
INNER JOIN Main_Table 
    ON Kogame_Data.Id_Kogame = Main_Table.Id_Kogame 
WHERE Main_Table.Current_OF = @var_OF;