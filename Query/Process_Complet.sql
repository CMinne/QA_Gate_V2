DECLARE @var_Main INT
DECLARE @var_OF VARCHAR(10)
DECLARE @var_Event INT
DECLARE @var_Reference VARCHAR(15)

DECLARE @var_Keyence_State BIT
DECLARE @var_Kogame_State BIT
DECLARE @var_Event_State BIT
DECLARE @var_Reference_State SMALLINT

SET @var_OF = '030621'
SET @var_Event = 0
SET @var_Keyence_State = 0
SET @var_Kogame_State = 0
SET @var_Event_State = 0
SET @var_Reference_State = 1

-- OK : 1 / NOK : 0
-- @var_Reference_State  0 : UE21 2000, 1 : UE21 2100, 2 : UE24 3200, 3 : UE24 3300

BEGIN TRANSACTION

-- ## Début de l'insertion ## --

IF(@var_Reference_State = 0)
	SELECT @var_Reference = nameReference 
	FROM QAGATE_1_Reference 
	WHERE idReference = 1
ELSE IF(@var_Reference_State = 1)
	SELECT @var_Reference = nameReference 
	FROM QAGATE_1_Reference 
	WHERE idReference = 2
ELSE IF(@var_Reference_State = 2)
	SELECT @var_Reference = nameReference 
	FROM QAGATE_1_Reference 
	WHERE idReference = 3
ELSE IF(@var_Reference_State = 3)
	SELECT @var_Reference = nameReference 
	FROM QAGATE_1_Reference 
	WHERE idReference = 4

INSERT INTO QAGATE_1_MainTable (reference     ,	currentOF) 
		VALUES				   (@var_Reference, @var_OF   )

SELECT @var_Main = MAX(QAGATE_1_MainTable.idPiece) 
FROM QAGATE_1_MainTable

INSERT INTO QAGATE_1_KeyenceData (currentOF, reference		)
VALUES							 (@Var_OF	, @var_Reference)

IF @var_Keyence_State = 0
BEGIN
		INSERT INTO Kogame_Data (Hauteur , Parallelisme_Face, Planeite_Face_Hexag, Planeite, Rectitude_Face_1, Rectitude_Face_2)
		VALUES				    (108.1872, 0.0298			, 0.0138			 , 0.0225  , 0.0049          , 0.0135		   )


		IF @var_Kogame_State = 0
		BEGIN
				UPDATE QAGATE_1_MainTable SET OK = 0 WHERE idPiece = @var_Main;
		END
		ELSE 
		BEGIN
			UPDATE QAGATE_1_MainTable SET kogameEtat = 1 WHERE idPiece = @var_Main;
		END
		
END
ELSE
BEGIN
	UPDATE QAGATE_1_MainTable SET keyenceEtat = 1 WHERE idPiece = @var_Main;
END

IF @var_Event_State = 1
BEGIN
		INSERT INTO QAGATE_1_EventData (reference	  , currentOF, code      , etat) 
		VALUES						   (@var_Reference, @var_OF	 , @var_Event, 2   )
END

COMMIT