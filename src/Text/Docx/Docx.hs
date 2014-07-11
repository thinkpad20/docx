{-# LANGUAGE LambdaCase, OverloadedStrings, RecordWildCards #-}
module Text.Docx.Docx where

import Prelude as P
import Control.Applicative
import Control.Monad.State
import Control.Monad.Identity
import Data.Monoid
import Data.ByteString as B
import Data.Maybe (isJust)
import Data.Sequence as Seq
import Data.Set as Set
import Data.HashMap.Strict as H
import Data.Hashable
import Data.String (IsString(..))
import Data.Text as T

type ImageData = ByteString
type ImageName = Text
data DocxMetaData = DocxMetaData
  { eRelations :: HashMap RelId Relation
  , eImages :: HashMap ImageName ImageData
  } deriving (Show)
data DocxDocument = DocxDocument
  { ddMetaData :: DocxMetaData
  , ddElement :: Element
  } deriving (Show)
data QName = QName (Maybe Text) Text deriving (Show)
instance IsString QName where
  fromString = w . T.pack
type Attribute = (QName, Text)
newtype FontFace = FontFace Text deriving (Show)
newtype FontSize = FontSize Int deriving (Show)
data TextConfig = TextConfig 
  { cfgFontFace :: Maybe FontFace
  , cfgFontSize :: Maybe FontSize} deriving (Show)
data Element = Element
  { eTag :: QName, eAttributes :: Seq Attribute, eText :: Text
  , eTail :: Text, eSubElements :: Seq Element
  } deriving (Show)

data TextStyle_ = Bold | Italic | Underline deriving (Show, Eq, Ord)
newtype TextStyle = TextStyle (Set TextStyle_) deriving (Show, Eq, Ord)

data Alignment = LeftAlign | RightAlign | CenterAlign deriving (Show)
data ParagraphConfig = ParagraphConfig
  { pcfgTextConfig :: TextConfig, pcfgStyle :: TextStyle
  , pcfgAlign :: Alignment}

newtype RelId = RelId Int deriving (Show, Eq, Ord)
instance Hashable RelId where
  hashWithSalt i (RelId j) = hashWithSalt i j

data RelType = ImageRelation
             | LinkRelation
             | StyleRelation
             | SettingsRelation
             deriving (Show)
type RelTarget = Text
data Relation = Relation RelType RelId RelTarget deriving (Show)

-- | A hyperlink.
type DisplayText = Text
type LinkUrl = Text
data Hyperlink = Hyperlink DisplayText LinkUrl deriving (Show)

-- | A piece of text, can be either a string or a hyperlink.
data TextItem = TextString Text
              | TextLink   Hyperlink
              deriving (Show)

-- | A chunk of text data, with one or more text items, and possible config.
data TextData = TextData (Maybe TextConfig) [TextItem]
              deriving (Show)

type Docx m a = StateT DocxMetaData m a
type ElementT m = StateT Element m

---------------------------------------------------------------------
-- Things which can be made into XML attributes get their own class.
---------------------------------------------------------------------

class Show a => AsAttribute a where toA :: a -> Text
                                    toA = T.pack . show

instance AsAttribute TextStyle_ where
  toA = \case Bold -> "b"; Italic -> "i"; Underline -> "u"
instance AsAttribute TextStyle where
  toA (TextStyle ts) = mconcat $ P.map toA $ Set.toList ts
instance AsAttribute FontSize
instance AsAttribute FontFace where toA (FontFace f) = f
instance AsAttribute Alignment
instance AsAttribute Text where toA = id
instance AsAttribute Int
instance AsAttribute RelId where toA (RelId i) = T.pack $ "rId" ++ show i

---------------------------------------------------------------------
-- Element builder monad
---------------------------------------------------------------------

-- | The element builder type.
type ElementBuilder m = ElementT m ()

-- | Builds an Element monadically, starting with the given tag.
buildElement :: Monad m => Text -> ElementBuilder m -> m Element
buildElement tag actions = execStateT actions $ docxElement tag

-- | Builds an element in the identity monad.
buildElementPure :: Text -> ElementBuilder Identity -> Element
buildElementPure tag = runIdentity . buildElement tag

-- | Sets the element tag.
setTag :: Monad m => QName -> ElementBuilder m
setTag t = modify $ \e -> e {eTag = t}

-- | Adds a subelement in the builder monad.
addSub :: Monad m => Element -> ElementBuilder m
addSub = modify . addSubelem

-- | Adds a subelement with the given tag and attributes in the builder monad.
addSub' :: Monad m => Text -> [Attribute] -> ElementBuilder m
addSub' t as = addSub (docxElement' t as)

---- | Builds a subelement and adds it in the builder monad.
buildSub :: Monad m => Text -> ElementBuilder m -> ElementBuilder m
buildSub tag actions = lift (buildElement tag actions) >>= addSub

-- | Adds an attribute in the builder monad.
addAttrib :: Monad m => Attribute -> ElementBuilder m
addAttrib attr = modify $ \e -> e {eAttributes = eAttributes e |> attr}

-- | Constructs a paragraph.
paragraph :: Monad m => ParagraphConfig -> Docx m Element
paragraph ParagraphConfig{..} = buildElement "p" $ do
  -- Build the paragraph properties
  buildSub "pPr" $ do
    addSub' "pStyle" [(w "val", toA pcfgStyle)]
    addSub' "jc" [(w "val", toA pcfgAlign)]
    let fsize = cfgFontSize pcfgTextConfig
        fface = cfgFontFace pcfgTextConfig
    -- If we have an fsize or fface, build an "rPr" element with the config
    when (isJust fsize || isJust fface) $ buildSub "rPr" $ do
      when (isJust fsize) $ do
        let Just size = toA <$> fsize
        addSub' "sz" [(w "val", size)]
        addSub' "szCs" [(w "val", size)]
      when (isJust fface) $ do
        let Just face = toA <$> fface
        addSub' "rFonts" [(w "ascii", face), (w "hAnsi", face)]

newRelationId :: Monad m => Docx m RelId
newRelationId = undefined

-- | Adds the first element as the last subelement of the second.
addSubelem :: Element -> Element -> Element
addSubelem subelem elem = elem {eSubElements = eSubElements elem |> subelem}

defaultElem :: Element
defaultElem = Element { eTag="", eText="", eTail="", eAttributes=mempty
                      , eSubElements=mempty }

-- | Appends text to the element.
appendText :: Text -> Element -> Element
appendText txt elem = 
  if T.length (eTail elem) > 0 then elem {eTail = eTail elem <> txt}
  else elem {eText = eText elem <> txt}

w :: Text -> QName
w = QName $ Just "w"

r :: Text -> QName
r = QName $ Just "r"

textElement :: Text -> Element
textElement content = (docxElement "t") {eText = content}

docxElement :: Text -> Element
docxElement tag = defaultElem {eTag = w tag}

docxElement' :: Text -> [(QName, Text)] -> Element
docxElement' tag attrs = (docxElement tag) {eAttributes = Seq.fromList attrs}

elemFromConfig :: Monad m => TextConfig -> Docx m Element
elemFromConfig = undefined

addRelation :: Monad m => Relation -> Docx m ()
addRelation = undefined

hyperlink :: Monad m => Hyperlink -> Docx m Element
hyperlink (Hyperlink content link) = buildElement "hyperlink" $ do
  addSub' "rStyle" [(w "val", "Hyperlink")]
  addSub $ textElement content
  relId <- lift newRelationId
  lift $ addRelation $ Relation LinkRelation relId link
  addAttrib (r "id", toA relId)

insertTextItem :: Monad m => TextItem -> Docx m Element
insertTextItem item = case item of
  TextString content -> return $ textElement content
  TextLink link -> hyperlink link

newDocument :: Monad m => m Element
newDocument = buildElement "document" $ addSub' "body" []

newDocumentPure = runIdentity newDocument

defaultMetaData = DocxMetaData mempty mempty

defaultDoc :: DocxDocument
defaultDoc = DocxDocument defaultMetaData defaultElem

--buildDocument :: Monad m => DocxBuilder m -> m DocxDocument
--buildDocument actions = execStateT actions $ defaultDoc

--insertTextData :: Monad m => TextData -> Docx m Element
--insertTextData (TextData mconfig items) = buildElement case mconfig of
--  Nothing -> go items where
--    go [] elem = return elem
--    go (item:items) elem = go items =<< insertTextItem item elem
--  Just cfg -> insertTextData (TextData Nothing items) =<< elemFromConfig cfg
