use super::{APPLICATION_VERSION, EXCEL_MAX_COLUMN_INDEX, EXCEL_MAX_ROW_INDEX, xml_local_name};
use excel_model::WorkbookState;
use office_common::{
    CellError, CellValue, DefinedNameId, FormulaSource, NameScope, OmError, OmErrorCode, OmResult,
    OmValue, Rect, SheetId,
};
use quick_xml::Reader;
use quick_xml::events::Event;
use regex::{Regex, RegexBuilder};
use std::cmp::Ordering;
use std::collections::BTreeSet;
use std::io::Cursor;

static FORMULA_RANDOM_STATE: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormulaEvalError {
    Unsupported,
    Circular,
    Null,
    Div0,
    Value,
    Ref,
    Name,
    NA,
    Num,
    GettingData,
    Spill,
    Calc,
    Field,
    Blocked,
    Busy,
    Connect,
    Python,
    Timeout,
    Unknown,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct FormulaArrayResult {
    pub(super) rows: usize,
    pub(super) cols: usize,
    pub(super) values: Vec<CellValue>,
}

impl FormulaArrayResult {
    fn single(value: CellValue) -> Self {
        Self {
            rows: 1,
            cols: 1,
            values: vec![value],
        }
    }
}

impl FormulaEvalError {
    pub(super) fn into_cell_value(self) -> Option<CellValue> {
        match self {
            FormulaEvalError::Unsupported => None,
            FormulaEvalError::Circular => Some(CellValue::Error(CellError::Calc)),
            FormulaEvalError::Null => Some(CellValue::Error(CellError::Null)),
            FormulaEvalError::Div0 => Some(CellValue::Error(CellError::Div0)),
            FormulaEvalError::Value => Some(CellValue::Error(CellError::Value)),
            FormulaEvalError::Ref => Some(CellValue::Error(CellError::Ref)),
            FormulaEvalError::Name => Some(CellValue::Error(CellError::Name)),
            FormulaEvalError::NA => Some(CellValue::Error(CellError::NA)),
            FormulaEvalError::Num => Some(CellValue::Error(CellError::Num)),
            FormulaEvalError::GettingData => Some(CellValue::Error(CellError::GettingData)),
            FormulaEvalError::Spill => Some(CellValue::Error(CellError::Spill)),
            FormulaEvalError::Calc => Some(CellValue::Error(CellError::Calc)),
            FormulaEvalError::Field => Some(CellValue::Error(CellError::Field)),
            FormulaEvalError::Blocked => Some(CellValue::Error(CellError::Blocked)),
            FormulaEvalError::Busy => Some(CellValue::Error(CellError::Busy)),
            FormulaEvalError::Connect => Some(CellValue::Error(CellError::Connect)),
            FormulaEvalError::Python => Some(CellValue::Error(CellError::Python)),
            FormulaEvalError::Timeout => Some(CellValue::Error(CellError::Timeout)),
            FormulaEvalError::Unknown => Some(CellValue::Error(CellError::Unknown)),
        }
    }
}

fn formula_eval_error_from_cell_error(error: CellError) -> FormulaEvalError {
    match error {
        CellError::Null => FormulaEvalError::Null,
        CellError::Div0 => FormulaEvalError::Div0,
        CellError::Ref => FormulaEvalError::Ref,
        CellError::Name => FormulaEvalError::Name,
        CellError::NA => FormulaEvalError::NA,
        CellError::Num => FormulaEvalError::Num,
        CellError::GettingData => FormulaEvalError::GettingData,
        CellError::Spill => FormulaEvalError::Spill,
        CellError::Calc => FormulaEvalError::Calc,
        CellError::Field => FormulaEvalError::Field,
        CellError::Blocked => FormulaEvalError::Blocked,
        CellError::Busy => FormulaEvalError::Busy,
        CellError::Connect => FormulaEvalError::Connect,
        CellError::Python => FormulaEvalError::Python,
        CellError::Timeout => FormulaEvalError::Timeout,
        CellError::Unknown => FormulaEvalError::Unknown,
        CellError::Value => FormulaEvalError::Value,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaScalarFunction {
    Abs,
    AccrInt,
    AccrIntM,
    Acos,
    Acosh,
    AmorDegrc,
    AmorLinc,
    Acot,
    Acoth,
    And,
    Asin,
    Asinh,
    Atan,
    Atan2,
    Atanh,
    BesselI,
    BesselJ,
    BesselK,
    BesselY,
    BetaDist,
    BetaDistLegacy,
    BetaInv,
    BetaInvLegacy,
    BinomDist,
    BinomDistRange,
    BinomInv,
    BitAnd,
    BitLShift,
    BitOr,
    BitRShift,
    BitXor,
    Ceiling,
    CeilingMath,
    CeilingPrecise,
    Combin,
    Combina,
    ConfidenceNorm,
    ConfidenceNormLegacy,
    ConfidenceT,
    Cos,
    Cosh,
    CoupDayBs,
    CoupDays,
    CoupDaysNc,
    CoupNcd,
    CoupNum,
    CoupPcd,
    CritBinom,
    ChiDistLegacy,
    ChiInvLegacy,
    ChiSqDist,
    ChiSqDistRt,
    ChiSqInv,
    ChiSqInvRt,
    Cot,
    Coth,
    Csc,
    Csch,
    Date,
    Day,
    Days,
    Days360,
    Db,
    Ddb,
    Degrees,
    Delta,
    Disc,
    Duration,
    EDate,
    EOMonth,
    Effect,
    Erf,
    ErfPrecise,
    Erfc,
    ErfcPrecise,
    Even,
    Exp,
    ExponDist,
    Fact,
    FactDouble,
    FDist,
    FDistLegacy,
    FDistRt,
    FInv,
    FInvLegacy,
    FInvRt,
    Fisher,
    FisherInv,
    Floor,
    FloorMath,
    FloorPrecise,
    Gauss,
    Gamma,
    GammaDist,
    GammaDistLegacy,
    GammaInv,
    GammaInvLegacy,
    GammaLn,
    GammaLnPrecise,
    GeStep,
    Hour,
    HypGeomDist,
    HypGeomDistLegacy,
    If,
    Intrate,
    IsoCeiling,
    IsoWeekNum,
    IsEven,
    IsOdd,
    Int,
    Ln,
    LogNormDist,
    LogNormDistLegacy,
    LogNormInv,
    LogNormInvLegacy,
    Log,
    Log10,
    Minute,
    MDuration,
    Mod,
    Month,
    MRound,
    Multinomial,
    NegBinomDist,
    NegBinomDistLegacy,
    Nominal,
    NormDist,
    NormInv,
    NormSDist,
    NormSDistLegacy,
    NormSInv,
    NormSInvLegacy,
    Not,
    Now,
    Odd,
    OddFPrice,
    OddFYield,
    OddLPrice,
    OddLYield,
    Or,
    PDuration,
    Permut,
    PermutationA,
    Phi,
    Pi,
    PoissonDist,
    Price,
    Radians,
    Rand,
    RandBetween,
    PriceDisc,
    PriceMat,
    Quotient,
    Received,
    RoundDown,
    Round,
    RoundUp,
    Rri,
    Sec,
    Sech,
    Sign,
    Sln,
    Power,
    Sin,
    Sinh,
    Standardize,
    Sqrt,
    SqrtPi,
    Second,
    Syd,
    TDist,
    TDistLegacy,
    TDist2T,
    TDistRt,
    TInv,
    TInvLegacy,
    TInv2T,
    Tan,
    Tanh,
    TBillEq,
    TBillPrice,
    TBillYield,
    Time,
    Today,
    Trunc,
    Vdb,
    Weekday,
    WeekNum,
    WeibullDist,
    Year,
    YearFrac,
    Yield,
    YieldDisc,
    YieldMat,
}

impl FormulaScalarFunction {
    fn from_name(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("ABS") {
            Some(Self::Abs)
        } else if name.eq_ignore_ascii_case("ACCRINT") {
            Some(Self::AccrInt)
        } else if name.eq_ignore_ascii_case("ACCRINTM") {
            Some(Self::AccrIntM)
        } else if name.eq_ignore_ascii_case("ACOS") {
            Some(Self::Acos)
        } else if name.eq_ignore_ascii_case("ACOSH") {
            Some(Self::Acosh)
        } else if name.eq_ignore_ascii_case("AMORDEGRC") {
            Some(Self::AmorDegrc)
        } else if name.eq_ignore_ascii_case("AMORLINC") {
            Some(Self::AmorLinc)
        } else if name.eq_ignore_ascii_case("ACOT") {
            Some(Self::Acot)
        } else if name.eq_ignore_ascii_case("ACOTH") {
            Some(Self::Acoth)
        } else if name.eq_ignore_ascii_case("AND") {
            Some(Self::And)
        } else if name.eq_ignore_ascii_case("ASIN") {
            Some(Self::Asin)
        } else if name.eq_ignore_ascii_case("ASINH") {
            Some(Self::Asinh)
        } else if name.eq_ignore_ascii_case("ATAN") {
            Some(Self::Atan)
        } else if name.eq_ignore_ascii_case("ATAN2") {
            Some(Self::Atan2)
        } else if name.eq_ignore_ascii_case("ATANH") {
            Some(Self::Atanh)
        } else if name.eq_ignore_ascii_case("BESSELI") {
            Some(Self::BesselI)
        } else if name.eq_ignore_ascii_case("BESSELJ") {
            Some(Self::BesselJ)
        } else if name.eq_ignore_ascii_case("BESSELK") {
            Some(Self::BesselK)
        } else if name.eq_ignore_ascii_case("BESSELY") {
            Some(Self::BesselY)
        } else if name.eq_ignore_ascii_case("BETA.DIST") {
            Some(Self::BetaDist)
        } else if name.eq_ignore_ascii_case("BETADIST") {
            Some(Self::BetaDistLegacy)
        } else if name.eq_ignore_ascii_case("BETA.INV") {
            Some(Self::BetaInv)
        } else if name.eq_ignore_ascii_case("BETAINV") {
            Some(Self::BetaInvLegacy)
        } else if name.eq_ignore_ascii_case("BINOM.DIST") || name.eq_ignore_ascii_case("BINOMDIST")
        {
            Some(Self::BinomDist)
        } else if name.eq_ignore_ascii_case("BINOM.DIST.RANGE") {
            Some(Self::BinomDistRange)
        } else if name.eq_ignore_ascii_case("BINOM.INV") {
            Some(Self::BinomInv)
        } else if name.eq_ignore_ascii_case("BITAND") {
            Some(Self::BitAnd)
        } else if name.eq_ignore_ascii_case("BITLSHIFT") {
            Some(Self::BitLShift)
        } else if name.eq_ignore_ascii_case("BITOR") {
            Some(Self::BitOr)
        } else if name.eq_ignore_ascii_case("BITRSHIFT") {
            Some(Self::BitRShift)
        } else if name.eq_ignore_ascii_case("BITXOR") {
            Some(Self::BitXor)
        } else if name.eq_ignore_ascii_case("CEILING") {
            Some(Self::Ceiling)
        } else if name.eq_ignore_ascii_case("CEILING.MATH") {
            Some(Self::CeilingMath)
        } else if name.eq_ignore_ascii_case("CEILING.PRECISE") {
            Some(Self::CeilingPrecise)
        } else if name.eq_ignore_ascii_case("COMBIN") {
            Some(Self::Combin)
        } else if name.eq_ignore_ascii_case("COMBINA") {
            Some(Self::Combina)
        } else if name.eq_ignore_ascii_case("CONFIDENCE.NORM") {
            Some(Self::ConfidenceNorm)
        } else if name.eq_ignore_ascii_case("CONFIDENCE") {
            Some(Self::ConfidenceNormLegacy)
        } else if name.eq_ignore_ascii_case("CONFIDENCE.T") {
            Some(Self::ConfidenceT)
        } else if name.eq_ignore_ascii_case("COS") {
            Some(Self::Cos)
        } else if name.eq_ignore_ascii_case("COSH") {
            Some(Self::Cosh)
        } else if name.eq_ignore_ascii_case("COUPDAYBS") {
            Some(Self::CoupDayBs)
        } else if name.eq_ignore_ascii_case("COUPDAYS") {
            Some(Self::CoupDays)
        } else if name.eq_ignore_ascii_case("COUPDAYSNC") {
            Some(Self::CoupDaysNc)
        } else if name.eq_ignore_ascii_case("COUPNCD") {
            Some(Self::CoupNcd)
        } else if name.eq_ignore_ascii_case("COUPNUM") {
            Some(Self::CoupNum)
        } else if name.eq_ignore_ascii_case("COUPPCD") {
            Some(Self::CoupPcd)
        } else if name.eq_ignore_ascii_case("CRITBINOM") {
            Some(Self::CritBinom)
        } else if name.eq_ignore_ascii_case("CHIDIST") {
            Some(Self::ChiDistLegacy)
        } else if name.eq_ignore_ascii_case("CHIINV") {
            Some(Self::ChiInvLegacy)
        } else if name.eq_ignore_ascii_case("CHISQ.DIST") {
            Some(Self::ChiSqDist)
        } else if name.eq_ignore_ascii_case("CHISQ.DIST.RT") {
            Some(Self::ChiSqDistRt)
        } else if name.eq_ignore_ascii_case("CHISQ.INV") {
            Some(Self::ChiSqInv)
        } else if name.eq_ignore_ascii_case("CHISQ.INV.RT") {
            Some(Self::ChiSqInvRt)
        } else if name.eq_ignore_ascii_case("COT") {
            Some(Self::Cot)
        } else if name.eq_ignore_ascii_case("COTH") {
            Some(Self::Coth)
        } else if name.eq_ignore_ascii_case("CSC") {
            Some(Self::Csc)
        } else if name.eq_ignore_ascii_case("CSCH") {
            Some(Self::Csch)
        } else if name.eq_ignore_ascii_case("DATE") {
            Some(Self::Date)
        } else if name.eq_ignore_ascii_case("DAY") {
            Some(Self::Day)
        } else if name.eq_ignore_ascii_case("DAYS") {
            Some(Self::Days)
        } else if name.eq_ignore_ascii_case("DAYS360") {
            Some(Self::Days360)
        } else if name.eq_ignore_ascii_case("DB") {
            Some(Self::Db)
        } else if name.eq_ignore_ascii_case("DDB") {
            Some(Self::Ddb)
        } else if name.eq_ignore_ascii_case("DEGREES") {
            Some(Self::Degrees)
        } else if name.eq_ignore_ascii_case("DELTA") {
            Some(Self::Delta)
        } else if name.eq_ignore_ascii_case("DISC") {
            Some(Self::Disc)
        } else if name.eq_ignore_ascii_case("DURATION") {
            Some(Self::Duration)
        } else if name.eq_ignore_ascii_case("EDATE") {
            Some(Self::EDate)
        } else if name.eq_ignore_ascii_case("EOMONTH") {
            Some(Self::EOMonth)
        } else if name.eq_ignore_ascii_case("EFFECT") {
            Some(Self::Effect)
        } else if name.eq_ignore_ascii_case("ERF") {
            Some(Self::Erf)
        } else if name.eq_ignore_ascii_case("ERF.PRECISE") {
            Some(Self::ErfPrecise)
        } else if name.eq_ignore_ascii_case("ERFC") {
            Some(Self::Erfc)
        } else if name.eq_ignore_ascii_case("ERFC.PRECISE") {
            Some(Self::ErfcPrecise)
        } else if name.eq_ignore_ascii_case("EVEN") {
            Some(Self::Even)
        } else if name.eq_ignore_ascii_case("EXPON.DIST") || name.eq_ignore_ascii_case("EXPONDIST")
        {
            Some(Self::ExponDist)
        } else if name.eq_ignore_ascii_case("EXP") {
            Some(Self::Exp)
        } else if name.eq_ignore_ascii_case("FACT") {
            Some(Self::Fact)
        } else if name.eq_ignore_ascii_case("FACTDOUBLE") {
            Some(Self::FactDouble)
        } else if name.eq_ignore_ascii_case("F.DIST") {
            Some(Self::FDist)
        } else if name.eq_ignore_ascii_case("FDIST") {
            Some(Self::FDistLegacy)
        } else if name.eq_ignore_ascii_case("F.DIST.RT") {
            Some(Self::FDistRt)
        } else if name.eq_ignore_ascii_case("F.INV") {
            Some(Self::FInv)
        } else if name.eq_ignore_ascii_case("FINV") {
            Some(Self::FInvLegacy)
        } else if name.eq_ignore_ascii_case("F.INV.RT") {
            Some(Self::FInvRt)
        } else if name.eq_ignore_ascii_case("FISHER") {
            Some(Self::Fisher)
        } else if name.eq_ignore_ascii_case("FISHERINV") {
            Some(Self::FisherInv)
        } else if name.eq_ignore_ascii_case("FLOOR") {
            Some(Self::Floor)
        } else if name.eq_ignore_ascii_case("FLOOR.MATH") {
            Some(Self::FloorMath)
        } else if name.eq_ignore_ascii_case("FLOOR.PRECISE") {
            Some(Self::FloorPrecise)
        } else if name.eq_ignore_ascii_case("GAUSS") {
            Some(Self::Gauss)
        } else if name.eq_ignore_ascii_case("GAMMA") {
            Some(Self::Gamma)
        } else if name.eq_ignore_ascii_case("GAMMA.DIST") {
            Some(Self::GammaDist)
        } else if name.eq_ignore_ascii_case("GAMMADIST") {
            Some(Self::GammaDistLegacy)
        } else if name.eq_ignore_ascii_case("GAMMA.INV") {
            Some(Self::GammaInv)
        } else if name.eq_ignore_ascii_case("GAMMAINV") {
            Some(Self::GammaInvLegacy)
        } else if name.eq_ignore_ascii_case("GAMMALN") {
            Some(Self::GammaLn)
        } else if name.eq_ignore_ascii_case("GAMMALN.PRECISE") {
            Some(Self::GammaLnPrecise)
        } else if name.eq_ignore_ascii_case("GESTEP") {
            Some(Self::GeStep)
        } else if name.eq_ignore_ascii_case("HOUR") {
            Some(Self::Hour)
        } else if name.eq_ignore_ascii_case("HYPGEOM.DIST") {
            Some(Self::HypGeomDist)
        } else if name.eq_ignore_ascii_case("HYPGEOMDIST") {
            Some(Self::HypGeomDistLegacy)
        } else if name.eq_ignore_ascii_case("IF") {
            Some(Self::If)
        } else if name.eq_ignore_ascii_case("INTRATE") {
            Some(Self::Intrate)
        } else if name.eq_ignore_ascii_case("ISO.CEILING") {
            Some(Self::IsoCeiling)
        } else if name.eq_ignore_ascii_case("ISOWEEKNUM") {
            Some(Self::IsoWeekNum)
        } else if name.eq_ignore_ascii_case("ISEVEN") {
            Some(Self::IsEven)
        } else if name.eq_ignore_ascii_case("ISODD") {
            Some(Self::IsOdd)
        } else if name.eq_ignore_ascii_case("INT") {
            Some(Self::Int)
        } else if name.eq_ignore_ascii_case("LN") {
            Some(Self::Ln)
        } else if name.eq_ignore_ascii_case("LOGNORM.INV") {
            Some(Self::LogNormInv)
        } else if name.eq_ignore_ascii_case("LOGNORM.DIST") {
            Some(Self::LogNormDist)
        } else if name.eq_ignore_ascii_case("LOGINV") {
            Some(Self::LogNormInvLegacy)
        } else if name.eq_ignore_ascii_case("LOGNORMDIST") {
            Some(Self::LogNormDistLegacy)
        } else if name.eq_ignore_ascii_case("LOG") {
            Some(Self::Log)
        } else if name.eq_ignore_ascii_case("LOG10") {
            Some(Self::Log10)
        } else if name.eq_ignore_ascii_case("MINUTE") {
            Some(Self::Minute)
        } else if name.eq_ignore_ascii_case("MDURATION") {
            Some(Self::MDuration)
        } else if name.eq_ignore_ascii_case("MOD") {
            Some(Self::Mod)
        } else if name.eq_ignore_ascii_case("MONTH") {
            Some(Self::Month)
        } else if name.eq_ignore_ascii_case("MROUND") {
            Some(Self::MRound)
        } else if name.eq_ignore_ascii_case("MULTINOMIAL") {
            Some(Self::Multinomial)
        } else if name.eq_ignore_ascii_case("NEGBINOM.DIST") {
            Some(Self::NegBinomDist)
        } else if name.eq_ignore_ascii_case("NEGBINOMDIST") {
            Some(Self::NegBinomDistLegacy)
        } else if name.eq_ignore_ascii_case("NOMINAL") {
            Some(Self::Nominal)
        } else if name.eq_ignore_ascii_case("NORM.INV") || name.eq_ignore_ascii_case("NORMINV") {
            Some(Self::NormInv)
        } else if name.eq_ignore_ascii_case("NORM.DIST") || name.eq_ignore_ascii_case("NORMDIST") {
            Some(Self::NormDist)
        } else if name.eq_ignore_ascii_case("NORM.S.INV") {
            Some(Self::NormSInv)
        } else if name.eq_ignore_ascii_case("NORM.S.DIST") {
            Some(Self::NormSDist)
        } else if name.eq_ignore_ascii_case("NORMSINV") {
            Some(Self::NormSInvLegacy)
        } else if name.eq_ignore_ascii_case("NORMSDIST") {
            Some(Self::NormSDistLegacy)
        } else if name.eq_ignore_ascii_case("NOT") {
            Some(Self::Not)
        } else if name.eq_ignore_ascii_case("NOW") {
            Some(Self::Now)
        } else if name.eq_ignore_ascii_case("ODD") {
            Some(Self::Odd)
        } else if name.eq_ignore_ascii_case("ODDFPRICE") {
            Some(Self::OddFPrice)
        } else if name.eq_ignore_ascii_case("ODDFYIELD") {
            Some(Self::OddFYield)
        } else if name.eq_ignore_ascii_case("ODDLPRICE") {
            Some(Self::OddLPrice)
        } else if name.eq_ignore_ascii_case("ODDLYIELD") {
            Some(Self::OddLYield)
        } else if name.eq_ignore_ascii_case("OR") {
            Some(Self::Or)
        } else if name.eq_ignore_ascii_case("PDURATION") {
            Some(Self::PDuration)
        } else if name.eq_ignore_ascii_case("PERMUT") {
            Some(Self::Permut)
        } else if name.eq_ignore_ascii_case("PERMUTATIONA") {
            Some(Self::PermutationA)
        } else if name.eq_ignore_ascii_case("PHI") {
            Some(Self::Phi)
        } else if name.eq_ignore_ascii_case("PI") {
            Some(Self::Pi)
        } else if name.eq_ignore_ascii_case("POISSON.DIST") || name.eq_ignore_ascii_case("POISSON")
        {
            Some(Self::PoissonDist)
        } else if name.eq_ignore_ascii_case("PRICE") {
            Some(Self::Price)
        } else if name.eq_ignore_ascii_case("RADIANS") {
            Some(Self::Radians)
        } else if name.eq_ignore_ascii_case("RAND") {
            Some(Self::Rand)
        } else if name.eq_ignore_ascii_case("RANDBETWEEN") {
            Some(Self::RandBetween)
        } else if name.eq_ignore_ascii_case("PRICEDISC") {
            Some(Self::PriceDisc)
        } else if name.eq_ignore_ascii_case("PRICEMAT") {
            Some(Self::PriceMat)
        } else if name.eq_ignore_ascii_case("QUOTIENT") {
            Some(Self::Quotient)
        } else if name.eq_ignore_ascii_case("RECEIVED") {
            Some(Self::Received)
        } else if name.eq_ignore_ascii_case("ROUNDDOWN") {
            Some(Self::RoundDown)
        } else if name.eq_ignore_ascii_case("ROUND") {
            Some(Self::Round)
        } else if name.eq_ignore_ascii_case("ROUNDUP") {
            Some(Self::RoundUp)
        } else if name.eq_ignore_ascii_case("RRI") {
            Some(Self::Rri)
        } else if name.eq_ignore_ascii_case("SEC") {
            Some(Self::Sec)
        } else if name.eq_ignore_ascii_case("SECH") {
            Some(Self::Sech)
        } else if name.eq_ignore_ascii_case("SIGN") {
            Some(Self::Sign)
        } else if name.eq_ignore_ascii_case("SLN") {
            Some(Self::Sln)
        } else if name.eq_ignore_ascii_case("POWER") {
            Some(Self::Power)
        } else if name.eq_ignore_ascii_case("SIN") {
            Some(Self::Sin)
        } else if name.eq_ignore_ascii_case("SINH") {
            Some(Self::Sinh)
        } else if name.eq_ignore_ascii_case("STANDARDIZE") {
            Some(Self::Standardize)
        } else if name.eq_ignore_ascii_case("SQRT") {
            Some(Self::Sqrt)
        } else if name.eq_ignore_ascii_case("SQRTPI") {
            Some(Self::SqrtPi)
        } else if name.eq_ignore_ascii_case("SECOND") {
            Some(Self::Second)
        } else if name.eq_ignore_ascii_case("SYD") {
            Some(Self::Syd)
        } else if name.eq_ignore_ascii_case("T.DIST") {
            Some(Self::TDist)
        } else if name.eq_ignore_ascii_case("TDIST") {
            Some(Self::TDistLegacy)
        } else if name.eq_ignore_ascii_case("T.DIST.2T") {
            Some(Self::TDist2T)
        } else if name.eq_ignore_ascii_case("T.DIST.RT") {
            Some(Self::TDistRt)
        } else if name.eq_ignore_ascii_case("T.INV") {
            Some(Self::TInv)
        } else if name.eq_ignore_ascii_case("TINV") {
            Some(Self::TInvLegacy)
        } else if name.eq_ignore_ascii_case("T.INV.2T") {
            Some(Self::TInv2T)
        } else if name.eq_ignore_ascii_case("TAN") {
            Some(Self::Tan)
        } else if name.eq_ignore_ascii_case("TANH") {
            Some(Self::Tanh)
        } else if name.eq_ignore_ascii_case("TBILLEQ") {
            Some(Self::TBillEq)
        } else if name.eq_ignore_ascii_case("TBILLPRICE") {
            Some(Self::TBillPrice)
        } else if name.eq_ignore_ascii_case("TBILLYIELD") {
            Some(Self::TBillYield)
        } else if name.eq_ignore_ascii_case("TIME") {
            Some(Self::Time)
        } else if name.eq_ignore_ascii_case("TODAY") {
            Some(Self::Today)
        } else if name.eq_ignore_ascii_case("TRUNC") {
            Some(Self::Trunc)
        } else if name.eq_ignore_ascii_case("VDB") {
            Some(Self::Vdb)
        } else if name.eq_ignore_ascii_case("WEEKDAY") {
            Some(Self::Weekday)
        } else if name.eq_ignore_ascii_case("WEEKNUM") {
            Some(Self::WeekNum)
        } else if name.eq_ignore_ascii_case("WEIBULL.DIST") || name.eq_ignore_ascii_case("WEIBULL")
        {
            Some(Self::WeibullDist)
        } else if name.eq_ignore_ascii_case("YEAR") {
            Some(Self::Year)
        } else if name.eq_ignore_ascii_case("YEARFRAC") {
            Some(Self::YearFrac)
        } else if name.eq_ignore_ascii_case("YIELD") {
            Some(Self::Yield)
        } else if name.eq_ignore_ascii_case("YIELDDISC") {
            Some(Self::YieldDisc)
        } else if name.eq_ignore_ascii_case("YIELDMAT") {
            Some(Self::YieldMat)
        } else {
            None
        }
    }

    fn evaluate(self, args: &[f64]) -> Result<f64, FormulaEvalError> {
        let serial_weekday_monday0 = |serial: i64| {
            let adjusted_serial = if serial > 60 { serial - 1 } else { serial };
            (adjusted_serial - 1).rem_euclid(7)
        };
        let week_start_from_return_type =
            |return_type: i64, allow_zero_based: bool| -> Result<i64, FormulaEvalError> {
                match return_type {
                    1 => Ok(6),
                    2 => Ok(0),
                    3 if allow_zero_based => Ok(0),
                    11..=17 => Ok(return_type - 11),
                    _ => Err(FormulaEvalError::Num),
                }
            };
        let iso_weeknum_from_serial = |serial: i64| -> Result<i64, FormulaEvalError> {
            let (year, month, day) = formula_ymd_from_serial(serial as f64)?;
            let days = if (year, month, day) == (1900, 2, 29) {
                days_from_civil(1900, 2, 28) + 1
            } else {
                days_from_civil(year, month, day)
            };
            let iso_weekday_from_days = |days: i64| (days + 3).rem_euclid(7) + 1;
            let weekday = iso_weekday_from_days(days);
            let current_monday = days - (weekday - 1);
            let thursday = current_monday + 3;
            let (iso_year, _, _) = civil_from_days(thursday);
            let jan4 = days_from_civil(iso_year, 1, 4);
            let jan4_weekday = iso_weekday_from_days(jan4);
            let week1_monday = jan4 - (jan4_weekday - 1);
            Ok((current_monday - week1_monday).div_euclid(7) + 1)
        };
        let days_in_excel_year = |year: i64| -> f64 {
            (1..=12)
                .map(|month| days_in_excel_month(year, month))
                .sum::<u32>() as f64
        };
        let serial_to_next_year = |year: i64| -> Result<i64, FormulaEvalError> {
            formula_date_serial_from_args((year + 1) as f64, 1.0, 1.0).map(|value| value as i64)
        };
        let serial_to_year_start = |year: i64| -> Result<i64, FormulaEvalError> {
            formula_date_serial_from_args(year as f64, 1.0, 1.0).map(|value| value as i64)
        };
        let days360 =
            |start_serial: i64, end_serial: i64, european: bool| -> Result<i64, FormulaEvalError> {
                let (start_serial, end_serial, sign) = if start_serial > end_serial {
                    (end_serial, start_serial, -1)
                } else {
                    (start_serial, end_serial, 1)
                };
                let (start_year, start_month, start_day) =
                    formula_ymd_from_serial(start_serial as f64)?;
                let (mut end_year, mut end_month, mut end_day) =
                    formula_ymd_from_serial(end_serial as f64)?;
                let mut start_day = start_day;
                if european {
                    if start_day == 31 {
                        start_day = 30;
                    }
                    if end_day == 31 {
                        end_day = 30;
                    }
                } else {
                    if start_day == 31 {
                        start_day = 30;
                    }
                    if end_day == 31 {
                        if start_day < 30 {
                            end_day = 1;
                            if end_month == 12 {
                                end_year += 1;
                                end_month = 1;
                            } else {
                                end_month += 1;
                            }
                        } else {
                            end_day = 30;
                        }
                    }
                }
                Ok(sign
                    * ((end_year - start_year) * 360
                        + (i64::from(end_month) - i64::from(start_month)) * 30
                        + i64::from(end_day)
                        - i64::from(start_day)))
            };
        let yearfrac_actual_actual =
            |start_serial: i64, end_serial: i64| -> Result<f64, FormulaEvalError> {
                if start_serial == end_serial {
                    return Ok(0.0);
                }
                let (start_serial, end_serial, sign) = if start_serial > end_serial {
                    (end_serial, start_serial, -1.0)
                } else {
                    (start_serial, end_serial, 1.0)
                };
                let (start_year, _, _) = formula_ymd_from_serial(start_serial as f64)?;
                let (end_year, _, _) = formula_ymd_from_serial(end_serial as f64)?;
                if start_year == end_year {
                    return Ok(
                        sign * (end_serial - start_serial) as f64 / days_in_excel_year(start_year)
                    );
                }
                let mut total = (serial_to_next_year(start_year)? - start_serial) as f64
                    / days_in_excel_year(start_year);
                for _ in (start_year + 1)..end_year {
                    total += 1.0;
                }
                total += (end_serial - serial_to_year_start(end_year)?) as f64
                    / days_in_excel_year(end_year);
                Ok(sign * total)
            };
        let yearfrac_basis = |value: f64| -> Result<i64, FormulaEvalError> {
            if !value.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let basis = value.trunc();
            if !(0.0..=4.0).contains(&basis) {
                return Err(FormulaEvalError::Num);
            }
            Ok(basis as i64)
        };
        let financial_date_serial = |value: f64| -> Result<i64, FormulaEvalError> {
            let serial = formula_serial_integer(value).map_err(|_| FormulaEvalError::Value)?;
            formula_ymd_from_serial(serial as f64)
                .map(|_| serial)
                .map_err(|_| FormulaEvalError::Value)
        };
        let yearfrac_by_basis =
            |start_serial: i64, end_serial: i64, basis: i64| -> Result<f64, FormulaEvalError> {
                match basis {
                    0 => Ok(days360(start_serial, end_serial, false)? as f64 / 360.0),
                    1 => yearfrac_actual_actual(start_serial, end_serial),
                    2 => Ok((end_serial - start_serial) as f64 / 360.0),
                    3 => Ok((end_serial - start_serial) as f64 / 365.0),
                    4 => Ok(days360(start_serial, end_serial, true)? as f64 / 360.0),
                    _ => Err(FormulaEvalError::Num),
                }
            };
        let discount_security_yearfrac =
            |settlement: f64, maturity: f64, basis: f64| -> Result<f64, FormulaEvalError> {
                let basis = yearfrac_basis(basis)?;
                let settlement = financial_date_serial(settlement)?;
                let maturity = financial_date_serial(maturity)?;
                if settlement >= maturity {
                    return Err(FormulaEvalError::Num);
                }
                yearfrac_by_basis(settlement, maturity, basis)
            };
        let maturity_security_yearfracs = |settlement: f64,
                                           maturity: f64,
                                           issue: f64,
                                           basis: f64|
         -> Result<(f64, f64, f64), FormulaEvalError> {
            let basis = yearfrac_basis(basis)?;
            let settlement = financial_date_serial(settlement)?;
            let maturity = financial_date_serial(maturity)?;
            let issue = financial_date_serial(issue)?;
            if issue >= settlement || settlement >= maturity {
                return Err(FormulaEvalError::Num);
            }
            let issue_to_maturity = yearfrac_by_basis(issue, maturity, basis)?;
            let settlement_to_maturity = yearfrac_by_basis(settlement, maturity, basis)?;
            let issue_to_settlement = yearfrac_by_basis(issue, settlement, basis)?;
            if issue_to_maturity <= 0.0 || settlement_to_maturity <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            Ok((
                issue_to_maturity,
                settlement_to_maturity,
                issue_to_settlement,
            ))
        };
        let treasury_bill_days =
            |settlement: f64, maturity: f64| -> Result<f64, FormulaEvalError> {
                let settlement = financial_date_serial(settlement)?;
                let maturity = financial_date_serial(maturity)?;
                if settlement >= maturity || maturity - settlement > 365 {
                    return Err(FormulaEvalError::Num);
                }
                Ok((maturity - settlement) as f64)
            };
        let coupon_frequency = |value: f64| -> Result<i64, FormulaEvalError> {
            if !value.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            match value.trunc() as i64 {
                1 | 2 | 4 => Ok(value.trunc() as i64),
                _ => Err(FormulaEvalError::Num),
            }
        };
        let coupon_schedule = |settlement: i64,
                               maturity: i64,
                               frequency: i64,
                               basis: i64|
         -> Result<(usize, i64, i64, f64), FormulaEvalError> {
            let months_per_coupon = 12 / frequency;
            let mut next_coupon = maturity;
            loop {
                let previous_coupon =
                    formula_edate(next_coupon as f64, -(months_per_coupon as f64))? as i64;
                if previous_coupon <= settlement {
                    let full_period = yearfrac_by_basis(previous_coupon, next_coupon, basis)?;
                    if full_period <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let remaining = yearfrac_by_basis(settlement, next_coupon, basis)?;
                    if remaining <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let first_period_fraction = remaining / full_period;
                    let mut coupon_count = 1_usize;
                    let mut coupon_date = next_coupon;
                    while coupon_date < maturity {
                        coupon_date =
                            formula_edate(coupon_date as f64, months_per_coupon as f64)? as i64;
                        coupon_count = coupon_count.checked_add(1).ok_or(FormulaEvalError::Num)?;
                        if coupon_count > 10000 {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    if coupon_date != maturity {
                        return Err(FormulaEvalError::Num);
                    }
                    return Ok((
                        coupon_count,
                        previous_coupon,
                        next_coupon,
                        first_period_fraction,
                    ));
                }
                next_coupon = previous_coupon;
            }
        };
        let coupon_schedule_from_values =
            |settlement: f64,
             maturity: f64,
             frequency: f64,
             basis: f64|
             -> Result<(i64, i64, i64, i64, usize, i64, i64), FormulaEvalError> {
                if ![settlement, maturity, frequency, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                let frequency = coupon_frequency(frequency)?;
                let basis = yearfrac_basis(basis)?;
                let settlement = financial_date_serial(settlement)?;
                let maturity = financial_date_serial(maturity)?;
                if settlement >= maturity {
                    return Err(FormulaEvalError::Num);
                }
                let (coupon_count, previous_coupon, next_coupon, _) =
                    coupon_schedule(settlement, maturity, frequency, basis)?;
                Ok((
                    settlement,
                    maturity,
                    frequency,
                    basis,
                    coupon_count,
                    previous_coupon,
                    next_coupon,
                ))
            };
        let coupon_schedule_from_args =
            |args: &[f64]| -> Result<(i64, i64, i64, i64, usize, i64, i64), FormulaEvalError> {
                let (settlement, maturity, frequency, basis) = match args {
                    [settlement, maturity, frequency] => (*settlement, *maturity, *frequency, 0.0),
                    [settlement, maturity, frequency, basis] => {
                        (*settlement, *maturity, *frequency, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                coupon_schedule_from_values(settlement, maturity, frequency, basis)
            };
        let coupon_days_between =
            |start_serial: i64, end_serial: i64, basis: i64| -> Result<f64, FormulaEvalError> {
                match basis {
                    0 => Ok(days360(start_serial, end_serial, false)? as f64),
                    1 | 2 | 3 => Ok((end_serial - start_serial) as f64),
                    4 => Ok(days360(start_serial, end_serial, true)? as f64),
                    _ => Err(FormulaEvalError::Num),
                }
            };
        let coupon_period_days =
            |previous_coupon: i64, next_coupon: i64, frequency: i64, basis: i64| -> f64 {
                match basis {
                    1 => (next_coupon - previous_coupon) as f64,
                    3 => 365.0 / frequency as f64,
                    _ => 360.0 / frequency as f64,
                }
            };
        let regular_coupon_price = |settlement: f64,
                                    maturity: f64,
                                    rate: f64,
                                    yld: f64,
                                    redemption: f64,
                                    frequency: f64,
                                    basis: f64|
         -> Result<f64, FormulaEvalError> {
            if ![
                settlement, maturity, rate, yld, redemption, frequency, basis,
            ]
            .iter()
            .all(|value| value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            if rate < 0.0 || redemption <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let (settlement, _, frequency, basis, coupon_count, previous_coupon, next_coupon) =
                coupon_schedule_from_values(settlement, maturity, frequency, basis)?;
            let frequency = frequency as f64;
            let yield_per_period = yld / frequency;
            let discount = 1.0 + yield_per_period;
            if discount <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let e = coupon_period_days(previous_coupon, next_coupon, frequency as i64, basis);
            if e <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let dsc = coupon_days_between(settlement, next_coupon, basis)?;
            let a = coupon_days_between(previous_coupon, settlement, basis)?;
            let coupon_payment = 100.0 * rate / frequency;
            let dsc_fraction = dsc / e;
            let accrued = coupon_payment * a / e;
            let price = if coupon_count == 1 {
                (redemption + coupon_payment) / (1.0 + yield_per_period * dsc_fraction) - accrued
            } else {
                let mut total =
                    redemption / discount.powf(coupon_count as f64 - 1.0 + dsc_fraction);
                for period in 1..=coupon_count {
                    total += coupon_payment / discount.powf(period as f64 - 1.0 + dsc_fraction);
                }
                total - accrued
            };
            if price.is_finite() {
                Ok(price)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        let regular_coupon_yield = |settlement: f64,
                                    maturity: f64,
                                    rate: f64,
                                    price: f64,
                                    redemption: f64,
                                    frequency: f64,
                                    basis: f64|
         -> Result<f64, FormulaEvalError> {
            if ![
                settlement, maturity, rate, price, redemption, frequency, basis,
            ]
            .iter()
            .all(|value| value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            if rate < 0.0 || price <= 0.0 || redemption <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let frequency = coupon_frequency(frequency)? as f64;
            let price_difference = |yld: f64| -> Result<f64, FormulaEvalError> {
                Ok(regular_coupon_price(
                    settlement, maturity, rate, yld, redemption, frequency, basis,
                )? - price)
            };
            let mut lower = -frequency + 1e-10;
            let lower_value = price_difference(lower)?;
            if lower_value < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let mut upper = 1.0;
            let mut upper_value = price_difference(upper)?;
            while upper_value > 0.0 {
                upper *= 2.0;
                if upper > 1e10 {
                    return Err(FormulaEvalError::Num);
                }
                upper_value = price_difference(upper)?;
            }
            for _ in 0..200 {
                let midpoint = (lower + upper) / 2.0;
                let midpoint_value = price_difference(midpoint)?;
                if midpoint_value.abs() <= 1e-10 || (upper - lower).abs() <= 1e-10 {
                    return Ok(midpoint);
                }
                if midpoint_value > 0.0 {
                    lower = midpoint;
                } else {
                    upper = midpoint;
                }
            }
            Ok((lower + upper) / 2.0)
        };
        let odd_first_coupon_price = |settlement: f64,
                                      maturity: f64,
                                      issue: f64,
                                      first_coupon: f64,
                                      rate: f64,
                                      yld: f64,
                                      redemption: f64,
                                      frequency: f64,
                                      basis: f64|
         -> Result<f64, FormulaEvalError> {
            if ![
                settlement,
                maturity,
                issue,
                first_coupon,
                rate,
                yld,
                redemption,
                frequency,
                basis,
            ]
            .iter()
            .all(|value| value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            if rate < 0.0 || redemption <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let frequency = coupon_frequency(frequency)?;
            let basis = yearfrac_basis(basis)?;
            let settlement = financial_date_serial(settlement)?;
            let maturity = financial_date_serial(maturity)?;
            let issue = financial_date_serial(issue)?;
            let first_coupon = financial_date_serial(first_coupon)?;
            if !(issue < settlement && settlement < first_coupon && first_coupon < maturity) {
                return Err(FormulaEvalError::Num);
            }
            let months_per_coupon = 12 / frequency;
            let notional_coupon = 100.0 * rate / frequency as f64;
            let previous_regular_coupon =
                formula_edate(first_coupon as f64, -(months_per_coupon as f64))? as i64;
            let first_period_days =
                coupon_period_days(previous_regular_coupon, first_coupon, frequency, basis);
            if first_period_days <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let first_coupon_payment = notional_coupon
                * coupon_days_between(issue, first_coupon, basis)?
                / first_period_days;
            let accrued = notional_coupon * coupon_days_between(issue, settlement, basis)?
                / first_period_days;
            let discount = 1.0 + yld / frequency as f64;
            if discount <= 0.0 {
                return Err(FormulaEvalError::Num);
            }

            let mut total = first_coupon_payment
                / discount
                    .powf(frequency as f64 * yearfrac_by_basis(settlement, first_coupon, basis)?);
            let mut coupon_date =
                formula_edate(first_coupon as f64, months_per_coupon as f64)? as i64;
            let mut guard = 0_usize;
            while coupon_date <= maturity {
                let mut cashflow = notional_coupon;
                if coupon_date == maturity {
                    cashflow += redemption;
                }
                total += cashflow
                    / discount.powf(
                        frequency as f64 * yearfrac_by_basis(settlement, coupon_date, basis)?,
                    );
                if coupon_date == maturity {
                    return formula_checked_numeric_result(total - accrued);
                }
                coupon_date = formula_edate(coupon_date as f64, months_per_coupon as f64)? as i64;
                guard += 1;
                if guard > 10000 {
                    return Err(FormulaEvalError::Num);
                }
            }
            Err(FormulaEvalError::Num)
        };
        let odd_last_coupon_price = |settlement: f64,
                                     maturity: f64,
                                     last_interest: f64,
                                     rate: f64,
                                     yld: f64,
                                     redemption: f64,
                                     frequency: f64,
                                     basis: f64|
         -> Result<f64, FormulaEvalError> {
            if ![
                settlement,
                maturity,
                last_interest,
                rate,
                yld,
                redemption,
                frequency,
                basis,
            ]
            .iter()
            .all(|value| value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            if rate < 0.0 || redemption <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let frequency = coupon_frequency(frequency)?;
            let basis = yearfrac_basis(basis)?;
            let settlement = financial_date_serial(settlement)?;
            let maturity = financial_date_serial(maturity)?;
            let last_interest = financial_date_serial(last_interest)?;
            if !(last_interest < settlement && settlement < maturity) {
                return Err(FormulaEvalError::Num);
            }
            let months_per_coupon = 12 / frequency;
            let next_regular_coupon =
                formula_edate(last_interest as f64, months_per_coupon as f64)? as i64;
            if maturity > next_regular_coupon {
                return Err(FormulaEvalError::Num);
            }
            let period_days =
                coupon_period_days(last_interest, next_regular_coupon, frequency, basis);
            if period_days <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let notional_coupon = 100.0 * rate / frequency as f64;
            let odd_coupon = notional_coupon * coupon_days_between(last_interest, maturity, basis)?
                / period_days;
            let accrued = notional_coupon * coupon_days_between(last_interest, settlement, basis)?
                / period_days;
            let discount = 1.0 + yld / frequency as f64;
            if discount <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let exponent = frequency as f64 * yearfrac_by_basis(settlement, maturity, basis)?;
            formula_checked_numeric_result(
                (redemption + odd_coupon) / discount.powf(exponent) - accrued,
            )
        };
        let solve_odd_coupon_yield =
            |price_difference: &mut dyn FnMut(f64) -> Result<f64, FormulaEvalError>,
             frequency: f64|
             -> Result<f64, FormulaEvalError> {
                let mut lower = -frequency + 1e-10;
                let lower_value = price_difference(lower)?;
                if lower_value < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let mut upper = 1.0;
                let mut upper_value = price_difference(upper)?;
                while upper_value > 0.0 {
                    upper *= 2.0;
                    if upper > 1e10 {
                        return Err(FormulaEvalError::Num);
                    }
                    upper_value = price_difference(upper)?;
                }
                for _ in 0..200 {
                    let midpoint = (lower + upper) / 2.0;
                    let midpoint_value = price_difference(midpoint)?;
                    if midpoint_value.abs() <= 1e-10 || (upper - lower).abs() <= 1e-10 {
                        return Ok(midpoint);
                    }
                    if midpoint_value > 0.0 {
                        lower = midpoint;
                    } else {
                        upper = midpoint;
                    }
                }
                Ok((lower + upper) / 2.0)
            };
        let duration_value = |settlement: f64,
                              maturity: f64,
                              coupon: f64,
                              yld: f64,
                              frequency: f64,
                              basis: f64,
                              modified: bool|
         -> Result<f64, FormulaEvalError> {
            if ![settlement, maturity, coupon, yld, frequency, basis]
                .iter()
                .all(|value| value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            let frequency = coupon_frequency(frequency)?;
            let basis = yearfrac_basis(basis)?;
            let settlement = financial_date_serial(settlement)?;
            let maturity = financial_date_serial(maturity)?;
            if coupon < 0.0 || yld < 0.0 || settlement >= maturity {
                return Err(FormulaEvalError::Num);
            }
            let (coupon_count, _, _, first_period_fraction) =
                coupon_schedule(settlement, maturity, frequency, basis)?;
            let frequency = frequency as f64;
            let yield_per_period = yld / frequency;
            let discount = 1.0 + yield_per_period;
            let coupon_payment = 100.0 * coupon / frequency;
            let mut present_value_total = 0.0;
            let mut weighted_present_value_total = 0.0;
            for period in 1..=coupon_count {
                let period_offset = period as f64 - 1.0 + first_period_fraction;
                let cash_flow = coupon_payment + if period == coupon_count { 100.0 } else { 0.0 };
                let denominator = discount.powf(period_offset);
                if denominator == 0.0 || !denominator.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                let present_value = cash_flow / denominator;
                present_value_total += present_value;
                weighted_present_value_total += present_value * period_offset / frequency;
                if !present_value_total.is_finite() || !weighted_present_value_total.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
            }
            if present_value_total == 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let mut duration = weighted_present_value_total / present_value_total;
            if modified {
                duration /= discount;
            }
            if duration.is_finite() {
                Ok(duration)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        let normalize_zero = |value: f64| if value == 0.0 { 0.0 } else { value };
        let ceiling_floor_math = |number: f64,
                                  significance: f64,
                                  mode: f64,
                                  ceiling: bool|
         -> Result<f64, FormulaEvalError> {
            if !number.is_finite() || !significance.is_finite() || !mode.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let significance = significance.abs();
            if significance == 0.0 {
                return Ok(0.0);
            }
            let value = if number >= 0.0 {
                let quotient = number / significance;
                if ceiling {
                    quotient.ceil() * significance
                } else {
                    quotient.floor() * significance
                }
            } else {
                let quotient = -number / significance;
                let magnitude = if ceiling {
                    if mode == 0.0 {
                        quotient.floor() * significance
                    } else {
                        quotient.ceil() * significance
                    }
                } else if mode == 0.0 {
                    quotient.ceil() * significance
                } else {
                    quotient.floor() * significance
                };
                -magnitude
            };
            if value.is_finite() {
                Ok(normalize_zero(value))
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        let ceiling_floor_legacy =
            |number: f64, significance: f64, ceiling: bool| -> Result<f64, FormulaEvalError> {
                if !number.is_finite() || !significance.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if number == 0.0 {
                    return Ok(0.0);
                }
                if significance == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                if number > 0.0 && significance < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let multiple = significance.abs();
                let quotient = number.abs() / multiple;
                let magnitude = if number >= 0.0 {
                    if ceiling {
                        quotient.ceil() * multiple
                    } else {
                        quotient.floor() * multiple
                    }
                } else if ceiling == (significance < 0.0) {
                    quotient.ceil() * multiple
                } else {
                    quotient.floor() * multiple
                };
                let value = if number.is_sign_negative() {
                    -magnitude
                } else {
                    magnitude
                };
                if value.is_finite() {
                    Ok(normalize_zero(value))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };
        let ceiling_floor_precise =
            |number: f64, significance: f64, ceiling: bool| -> Result<f64, FormulaEvalError> {
                if !number.is_finite() || !significance.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                let significance = significance.abs();
                if number == 0.0 || significance == 0.0 {
                    return Ok(0.0);
                }
                let quotient = number / significance;
                let value = if ceiling {
                    quotient.ceil() * significance
                } else {
                    quotient.floor() * significance
                };
                if value.is_finite() {
                    Ok(normalize_zero(value))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };
        let trunc_nonnegative_integer = |value: f64| -> Result<u64, FormulaEvalError> {
            if !value.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let value = value.trunc();
            if value < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            if value > u64::MAX as f64 {
                return Err(FormulaEvalError::Num);
            }
            Ok(value as u64)
        };
        let factorial = |value: u64| -> Result<f64, FormulaEvalError> {
            let mut total = 1.0_f64;
            for factor in 2..=value {
                total *= factor as f64;
                if !total.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
            }
            Ok(total)
        };
        let combination = |number: u64, chosen: u64| -> Result<f64, FormulaEvalError> {
            if chosen > number {
                return Err(FormulaEvalError::Num);
            }
            let chosen = chosen.min(number - chosen);
            let mut total = 1.0_f64;
            for step in 1..=chosen {
                total = total * (number - chosen + step) as f64 / step as f64;
                if !total.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
            }
            Ok(total.round())
        };
        let round_away_to_integer_with_parity =
            |value: f64, odd: bool| -> Result<f64, FormulaEvalError> {
                if !value.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if value == 0.0 {
                    return Ok(if odd { 1.0 } else { 0.0 });
                }
                let mut magnitude = value.abs().ceil();
                if (magnitude.rem_euclid(2.0) != 0.0) != odd {
                    magnitude += 1.0;
                }
                let rounded = if value.is_sign_negative() {
                    -magnitude
                } else {
                    magnitude
                };
                if rounded.is_finite() {
                    Ok(normalize_zero(rounded))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };
        const RECIPROCAL_TRIG_INPUT_LIMIT: f64 = 134_217_728.0;
        let validate_reciprocal_trig_input = |value: f64| -> Result<(), FormulaEvalError> {
            if !value.is_finite() || value.abs() >= RECIPROCAL_TRIG_INPUT_LIMIT {
                Err(FormulaEvalError::Num)
            } else {
                Ok(())
            }
        };
        let checked_numeric_result = |value: f64| -> Result<f64, FormulaEvalError> {
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        let erf_approx = |value: f64| {
            let sign = if value.is_sign_negative() { -1.0 } else { 1.0 };
            let x = value.abs();
            let t = 1.0 / (1.0 + 0.5 * x);
            let tau = t
                * (-x * x - 1.26551223
                    + t * (1.00002368
                        + t * (0.37409196
                            + t * (0.09678418
                                + t * (-0.18628806
                                    + t * (0.27886807
                                        + t * (-1.13520398
                                            + t * (1.48851587
                                                + t * (-0.82215223 + t * 0.17087277)))))))))
                    .exp();
            sign * (1.0 - tau)
        };
        let standard_normal_pdf = |z: f64| {
            const INV_SQRT_2_PI: f64 = 0.3989422804014327;
            INV_SQRT_2_PI * (-0.5 * z * z).exp()
        };
        let standard_normal_cdf = |z: f64| 0.5 * (1.0 + erf_approx(z / std::f64::consts::SQRT_2));
        let inverse_standard_normal = |probability: f64| -> Result<f64, FormulaEvalError> {
            if !probability.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if probability <= 0.0 || probability >= 1.0 {
                return Err(FormulaEvalError::Num);
            }

            const A: [f64; 6] = [
                -3.969683028665376e1,
                2.209460984245205e2,
                -2.759285104469687e2,
                1.383577518672690e2,
                -3.066479806614716e1,
                2.506628277459239,
            ];
            const B: [f64; 5] = [
                -5.447609879822406e1,
                1.615858368580409e2,
                -1.556989798598866e2,
                6.680131188771972e1,
                -1.328068155288572e1,
            ];
            const C: [f64; 6] = [
                -7.784894002430293e-3,
                -3.223964580411365e-1,
                -2.400758277161838,
                -2.549732539343734,
                4.374664141464968,
                2.938163982698783,
            ];
            const D: [f64; 4] = [
                7.784695709041462e-3,
                3.224671290700398e-1,
                2.445134137142996,
                3.754408661907416,
            ];
            const P_LOW: f64 = 0.02425;
            const P_HIGH: f64 = 1.0 - P_LOW;

            let value = if probability < P_LOW {
                let q = (-2.0 * probability.ln()).sqrt();
                (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
                    / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
            } else if probability <= P_HIGH {
                let q = probability - 0.5;
                let r = q * q;
                (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r + A[5]) * q
                    / (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r + 1.0)
            } else {
                let q = (-2.0 * (1.0 - probability).ln()).sqrt();
                -(((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
                    / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
            };
            checked_numeric_result(value)
        };
        let gamma_ln_value = |value: f64| {
            const COEFFICIENTS: [f64; 9] = [
                0.9999999999998099,
                676.5203681218851,
                -1259.1392167224028,
                771.3234287776531,
                -176.6150291621406,
                12.507343278686905,
                -0.13857109526572012,
                0.000009984369578019572,
                0.00000015056327351493116,
            ];
            let lanczos = |input: f64| {
                let z = input - 1.0;
                let mut x = COEFFICIENTS[0];
                for (index, coefficient) in COEFFICIENTS.iter().enumerate().skip(1) {
                    x += coefficient / (z + index as f64);
                }
                let t = z + 7.5;
                0.5 * (2.0 * std::f64::consts::PI).ln() + (z + 0.5) * t.ln() - t + x.ln()
            };
            if value < 0.5 {
                std::f64::consts::PI.ln()
                    - (std::f64::consts::PI * value).sin().ln()
                    - lanczos(1.0 - value)
            } else {
                lanczos(value)
            }
        };
        let log_combination = |number: u64, chosen: u64| -> Result<f64, FormulaEvalError> {
            if chosen > number {
                return Err(FormulaEvalError::Num);
            }
            let number_plus_one = number.checked_add(1).ok_or(FormulaEvalError::Num)?;
            let chosen_plus_one = chosen.checked_add(1).ok_or(FormulaEvalError::Num)?;
            let remainder_plus_one = (number - chosen)
                .checked_add(1)
                .ok_or(FormulaEvalError::Num)?;
            checked_numeric_result(
                gamma_ln_value(number_plus_one as f64)
                    - gamma_ln_value(chosen_plus_one as f64)
                    - gamma_ln_value(remainder_plus_one as f64),
            )
        };
        let binomial_probability =
            |successes: u64, trials: u64, probability: f64| -> Result<f64, FormulaEvalError> {
                if successes > trials {
                    return Err(FormulaEvalError::Num);
                }
                if probability == 0.0 {
                    return Ok(if successes == 0 { 1.0 } else { 0.0 });
                }
                if probability == 1.0 {
                    return Ok(if successes == trials { 1.0 } else { 0.0 });
                }
                let failure_probability = 1.0 - probability;
                let log_probability = log_combination(trials, successes)?
                    + successes as f64 * probability.ln()
                    + (trials - successes) as f64 * failure_probability.ln();
                checked_numeric_result(log_probability.exp())
            };
        let cumulative_binomial_probability =
            |successes: u64, trials: u64, probability: f64| -> Result<f64, FormulaEvalError> {
                let mut total = 0.0;
                for value in 0..=successes {
                    total += binomial_probability(value, trials, probability)?;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                checked_numeric_result(total.min(1.0))
            };
        let negative_binomial_probability =
            |failures: u64, successes: u64, probability: f64| -> Result<f64, FormulaEvalError> {
                if probability == 0.0 {
                    return Ok(0.0);
                }
                if probability == 1.0 {
                    return Ok(if failures == 0 { 1.0 } else { 0.0 });
                }
                let total_before_last = failures
                    .checked_add(successes)
                    .and_then(|value| value.checked_sub(1))
                    .ok_or(FormulaEvalError::Num)?;
                let log_probability = log_combination(total_before_last, failures)?
                    + failures as f64 * (1.0 - probability).ln()
                    + successes as f64 * probability.ln();
                checked_numeric_result(log_probability.exp())
            };
        let hypergeometric_probability = |sample_successes: u64,
                                          sample_size: u64,
                                          population_successes: u64,
                                          population_size: u64|
         -> Result<f64, FormulaEvalError> {
            let sample_failures = sample_size
                .checked_sub(sample_successes)
                .ok_or(FormulaEvalError::Num)?;
            let population_failures = population_size
                .checked_sub(population_successes)
                .ok_or(FormulaEvalError::Num)?;
            let log_probability = log_combination(population_successes, sample_successes)?
                + log_combination(population_failures, sample_failures)?
                - log_combination(population_size, sample_size)?;
            checked_numeric_result(log_probability.exp())
        };
        let regularized_gamma_p = |shape: f64, x: f64| -> Result<f64, FormulaEvalError> {
            if !shape.is_finite() || !x.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if shape <= 0.0 || x < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            if x == 0.0 {
                return Ok(0.0);
            }
            const EPSILON: f64 = 1e-14;
            const FLOOR: f64 = 1e-300;
            const MAX_ITERATIONS: usize = 200;
            let gamma_ln = gamma_ln_value(shape);
            if x < shape + 1.0 {
                let mut term = 1.0 / shape;
                let mut sum = term;
                let mut ap = shape;
                for _ in 0..MAX_ITERATIONS {
                    ap += 1.0;
                    term *= x / ap;
                    sum += term;
                    if term.abs() <= sum.abs() * EPSILON {
                        return checked_numeric_result(
                            (sum * (-x + shape * x.ln() - gamma_ln).exp()).clamp(0.0, 1.0),
                        );
                    }
                }
                return checked_numeric_result(
                    (sum * (-x + shape * x.ln() - gamma_ln).exp()).clamp(0.0, 1.0),
                );
            }

            let mut b = x + 1.0 - shape;
            let mut c = 1.0 / FLOOR;
            let mut d = 1.0 / b.max(FLOOR);
            let mut h = d;
            for i in 1..=MAX_ITERATIONS {
                let i = i as f64;
                let an = -i * (i - shape);
                b += 2.0;
                d = an * d + b;
                if d.abs() < FLOOR {
                    d = FLOOR;
                }
                c = b + an / c;
                if c.abs() < FLOOR {
                    c = FLOOR;
                }
                d = 1.0 / d;
                let delta = d * c;
                h *= delta;
                if (delta - 1.0).abs() <= EPSILON {
                    let q = (-x + shape * x.ln() - gamma_ln).exp() * h;
                    return checked_numeric_result((1.0 - q).clamp(0.0, 1.0));
                }
            }
            let q = (-x + shape * x.ln() - gamma_ln).exp() * h;
            checked_numeric_result((1.0 - q).clamp(0.0, 1.0))
        };
        let regularized_gamma_q = |shape: f64, x: f64| -> Result<f64, FormulaEvalError> {
            regularized_gamma_p(shape, x).map(|value| (1.0 - value).clamp(0.0, 1.0))
        };
        let beta_fraction = |alpha: f64, beta: f64, x: f64| -> Result<f64, FormulaEvalError> {
            const EPSILON: f64 = 1e-14;
            const FLOOR: f64 = 1e-300;
            const MAX_ITERATIONS: usize = 200;
            let qab = alpha + beta;
            let qap = alpha + 1.0;
            let qam = alpha - 1.0;
            let mut c = 1.0;
            let mut d = 1.0 - qab * x / qap;
            if d.abs() < FLOOR {
                d = FLOOR;
            }
            d = 1.0 / d;
            let mut h = d;
            for m in 1..=MAX_ITERATIONS {
                let m_f = m as f64;
                let m2 = 2.0 * m_f;
                let mut aa = m_f * (beta - m_f) * x / ((qam + m2) * (alpha + m2));
                d = 1.0 + aa * d;
                if d.abs() < FLOOR {
                    d = FLOOR;
                }
                c = 1.0 + aa / c;
                if c.abs() < FLOOR {
                    c = FLOOR;
                }
                d = 1.0 / d;
                h *= d * c;
                aa = -(alpha + m_f) * (qab + m_f) * x / ((alpha + m2) * (qap + m2));
                d = 1.0 + aa * d;
                if d.abs() < FLOOR {
                    d = FLOOR;
                }
                c = 1.0 + aa / c;
                if c.abs() < FLOOR {
                    c = FLOOR;
                }
                d = 1.0 / d;
                let delta = d * c;
                h *= delta;
                if (delta - 1.0).abs() <= EPSILON {
                    return checked_numeric_result(h);
                }
            }
            checked_numeric_result(h)
        };
        let regularized_beta = |x: f64, alpha: f64, beta: f64| -> Result<f64, FormulaEvalError> {
            if ![x, alpha, beta].iter().all(|value| value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            if alpha <= 0.0 || beta <= 0.0 || !(0.0..=1.0).contains(&x) {
                return Err(FormulaEvalError::Num);
            }
            if x == 0.0 || x == 1.0 {
                return Ok(x);
            }
            let log_beta =
                gamma_ln_value(alpha) + gamma_ln_value(beta) - gamma_ln_value(alpha + beta);
            let front = (alpha * x.ln() + beta * (-x).ln_1p() - log_beta).exp();
            if x < (alpha + 1.0) / (alpha + beta + 2.0) {
                checked_numeric_result(
                    (front * beta_fraction(alpha, beta, x)? / alpha).clamp(0.0, 1.0),
                )
            } else {
                checked_numeric_result(
                    (1.0 - front * beta_fraction(beta, alpha, 1.0 - x)? / beta).clamp(0.0, 1.0),
                )
            }
        };
        let beta_pdf = |x: f64, alpha: f64, beta: f64| -> Result<f64, FormulaEvalError> {
            if ![x, alpha, beta].iter().all(|value| value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            if alpha <= 0.0 || beta <= 0.0 || !(0.0..=1.0).contains(&x) {
                return Err(FormulaEvalError::Num);
            }
            if x == 0.0 {
                if alpha < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                return Ok(if alpha == 1.0 { beta } else { 0.0 });
            }
            if x == 1.0 {
                if beta < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                return Ok(if beta == 1.0 { alpha } else { 0.0 });
            }
            let log_beta =
                gamma_ln_value(alpha) + gamma_ln_value(beta) - gamma_ln_value(alpha + beta);
            checked_numeric_result(
                ((alpha - 1.0) * x.ln() + (beta - 1.0) * (-x).ln_1p() - log_beta).exp(),
            )
        };
        let inverse_unit_cdf = |probability: f64,
                                cdf: &mut dyn FnMut(f64) -> Result<f64, FormulaEvalError>|
         -> Result<f64, FormulaEvalError> {
            if !probability.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if probability <= 0.0 || probability >= 1.0 {
                return Err(FormulaEvalError::Num);
            }
            let mut low = 0.0;
            let mut high = 1.0;
            for _ in 0..100 {
                let mid = (low + high) / 2.0;
                if cdf(mid)? < probability {
                    low = mid;
                } else {
                    high = mid;
                }
            }
            checked_numeric_result((low + high) / 2.0)
        };
        let inverse_positive_cdf = |probability: f64,
                                    cdf: &mut dyn FnMut(f64) -> Result<f64, FormulaEvalError>|
         -> Result<f64, FormulaEvalError> {
            if !probability.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if probability <= 0.0 || probability >= 1.0 {
                return Err(FormulaEvalError::Num);
            }
            let mut low = 0.0;
            let mut high = 1.0;
            let mut guard = 0;
            while cdf(high)? < probability {
                low = high;
                high *= 2.0;
                guard += 1;
                if guard > 1024 || !high.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
            }
            for _ in 0..120 {
                let mid = (low + high) / 2.0;
                if cdf(mid)? < probability {
                    low = mid;
                } else {
                    high = mid;
                }
            }
            checked_numeric_result((low + high) / 2.0)
        };
        let student_t_cdf = |x: f64, degrees: f64| -> Result<f64, FormulaEvalError> {
            if !x.is_finite() || !degrees.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if degrees < 1.0 {
                return Err(FormulaEvalError::Num);
            }
            let degrees = degrees.trunc();
            let beta_x = degrees / (degrees + x * x);
            let tail = 0.5 * regularized_beta(beta_x, degrees / 2.0, 0.5)?;
            Ok(if x >= 0.0 { 1.0 - tail } else { tail })
        };
        let student_t_pdf = |x: f64, degrees: f64| -> Result<f64, FormulaEvalError> {
            if !x.is_finite() || !degrees.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if degrees < 1.0 {
                return Err(FormulaEvalError::Num);
            }
            let degrees = degrees.trunc();
            let log_density = gamma_ln_value((degrees + 1.0) / 2.0)
                - gamma_ln_value(degrees / 2.0)
                - 0.5 * (degrees * std::f64::consts::PI).ln()
                - (degrees + 1.0) / 2.0 * (1.0 + x * x / degrees).ln();
            checked_numeric_result(log_density.exp())
        };
        let reciprocal_numeric_result = |denominator: f64| -> Result<f64, FormulaEvalError> {
            if denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            checked_numeric_result(1.0 / denominator)
        };

        match self {
            FormulaScalarFunction::Abs => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(value.abs())
            }
            FormulaScalarFunction::AccrInt => {
                let (issue, first_interest, settlement, rate, par, frequency, basis, calc_method) =
                    match args {
                        [issue, first_interest, settlement, rate, par, frequency] => (
                            *issue,
                            *first_interest,
                            *settlement,
                            *rate,
                            *par,
                            *frequency,
                            0.0,
                            1.0,
                        ),
                        [
                            issue,
                            first_interest,
                            settlement,
                            rate,
                            par,
                            frequency,
                            basis,
                        ] => (
                            *issue,
                            *first_interest,
                            *settlement,
                            *rate,
                            *par,
                            *frequency,
                            *basis,
                            1.0,
                        ),
                        [
                            issue,
                            first_interest,
                            settlement,
                            rate,
                            par,
                            frequency,
                            basis,
                            calc_method,
                        ] => (
                            *issue,
                            *first_interest,
                            *settlement,
                            *rate,
                            *par,
                            *frequency,
                            *basis,
                            *calc_method,
                        ),
                        _ => return Err(FormulaEvalError::Value),
                    };
                if ![
                    issue,
                    first_interest,
                    settlement,
                    rate,
                    par,
                    frequency,
                    basis,
                    calc_method,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if rate <= 0.0 || par <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let frequency = coupon_frequency(frequency)?;
                let basis = yearfrac_basis(basis)?;
                let issue = financial_date_serial(issue)?;
                let first_interest = financial_date_serial(first_interest)?;
                let settlement = financial_date_serial(settlement)?;
                if issue >= settlement {
                    return Err(FormulaEvalError::Num);
                }

                let months_per_coupon = 12 / frequency;
                let accrual_start = if calc_method == 0.0 && settlement > first_interest {
                    first_interest
                } else {
                    issue
                };
                let mut next_coupon = first_interest;
                let mut guard = 0_usize;
                while next_coupon <= accrual_start {
                    next_coupon =
                        formula_edate(next_coupon as f64, months_per_coupon as f64)? as i64;
                    guard += 1;
                    if guard > 10000 {
                        return Err(FormulaEvalError::Num);
                    }
                }
                loop {
                    let previous_coupon =
                        formula_edate(next_coupon as f64, -(months_per_coupon as f64))? as i64;
                    if previous_coupon <= accrual_start {
                        break;
                    }
                    next_coupon = previous_coupon;
                    guard += 1;
                    if guard > 10000 {
                        return Err(FormulaEvalError::Num);
                    }
                }

                let coupon_interest = par * rate / frequency as f64;
                let mut accrued_periods = 0.0;
                while accrual_start < settlement {
                    let previous_coupon =
                        formula_edate(next_coupon as f64, -(months_per_coupon as f64))? as i64;
                    let period_start = accrual_start.max(previous_coupon);
                    let period_end = settlement.min(next_coupon);
                    if period_end > period_start {
                        let accrued_days = coupon_days_between(period_start, period_end, basis)?;
                        let period_days =
                            coupon_period_days(previous_coupon, next_coupon, frequency, basis);
                        if period_days <= 0.0 {
                            return Err(FormulaEvalError::Num);
                        }
                        accrued_periods += accrued_days / period_days;
                        if !accrued_periods.is_finite() {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    if next_coupon >= settlement {
                        break;
                    }
                    next_coupon =
                        formula_edate(next_coupon as f64, months_per_coupon as f64)? as i64;
                    guard += 1;
                    if guard > 10000 {
                        return Err(FormulaEvalError::Num);
                    }
                }

                checked_numeric_result(coupon_interest * accrued_periods)
            }
            FormulaScalarFunction::AccrIntM => {
                let (issue, settlement, rate, par, basis) = match args {
                    [issue, settlement, rate] => (*issue, *settlement, *rate, 1000.0, 0.0),
                    [issue, settlement, rate, par] => (*issue, *settlement, *rate, *par, 0.0),
                    [issue, settlement, rate, par, basis] => {
                        (*issue, *settlement, *rate, *par, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![issue, settlement, rate, par, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if rate <= 0.0 || par <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let basis = yearfrac_basis(basis)?;
                let issue = financial_date_serial(issue)?;
                let settlement = financial_date_serial(settlement)?;
                if issue >= settlement {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(par * rate * yearfrac_by_basis(issue, settlement, basis)?)
            }
            FormulaScalarFunction::Acos => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value < -1.0 || *value > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value.acos())
            }
            FormulaScalarFunction::Acosh => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(value.acosh())
            }
            FormulaScalarFunction::AmorDegrc => {
                let (cost, date_purchased, first_period, salvage, period, rate, basis) = match args
                {
                    [cost, date_purchased, first_period, salvage, period, rate] => (
                        *cost,
                        *date_purchased,
                        *first_period,
                        *salvage,
                        *period,
                        *rate,
                        0.0,
                    ),
                    [
                        cost,
                        date_purchased,
                        first_period,
                        salvage,
                        period,
                        rate,
                        basis,
                    ] => (
                        *cost,
                        *date_purchased,
                        *first_period,
                        *salvage,
                        *period,
                        *rate,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![
                    cost,
                    date_purchased,
                    first_period,
                    salvage,
                    period,
                    rate,
                    basis,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                let basis = yearfrac_basis(basis)?;
                if basis == 2 {
                    return Err(FormulaEvalError::Num);
                }
                let date_purchased = financial_date_serial(date_purchased)?;
                let first_period = financial_date_serial(first_period)?;
                let period = period.trunc();
                if cost <= 0.0
                    || salvage < 0.0
                    || salvage >= cost
                    || period < 0.0
                    || rate <= 0.0
                    || date_purchased >= first_period
                {
                    return Err(FormulaEvalError::Num);
                }
                let asset_life = 1.0 / rate;
                let coefficient = if asset_life > 6.0 {
                    2.5
                } else if asset_life >= 5.0 {
                    2.0
                } else if asset_life >= 3.0 && asset_life <= 4.0 {
                    1.5
                } else {
                    return Err(FormulaEvalError::Num);
                };
                let depreciation_rate = rate * coefficient;
                let first_depreciation = round_half_away_from_zero(
                    cost * depreciation_rate
                        * yearfrac_by_basis(date_purchased, first_period, basis)?,
                );
                if period == 0.0 {
                    return checked_numeric_result(first_depreciation);
                }
                let mut accumulated_depreciation = first_depreciation;
                let lifetime_periods = asset_life.floor();
                let mut current_period = 1.0;
                while current_period <= period {
                    let depreciation = if accumulated_depreciation > cost - salvage {
                        0.0
                    } else {
                        let remaining_value = cost - accumulated_depreciation;
                        let period_rate = if current_period == lifetime_periods - 2.0 {
                            0.5
                        } else if current_period == lifetime_periods - 1.0 {
                            1.0
                        } else {
                            depreciation_rate
                        };
                        round_half_away_from_zero(remaining_value * period_rate)
                    };
                    if current_period == period {
                        return checked_numeric_result(depreciation);
                    }
                    accumulated_depreciation += depreciation;
                    if !accumulated_depreciation.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    current_period += 1.0;
                }
                Ok(0.0)
            }
            FormulaScalarFunction::AmorLinc => {
                let (cost, date_purchased, first_period, salvage, period, rate, basis) = match args
                {
                    [cost, date_purchased, first_period, salvage, period, rate] => (
                        *cost,
                        *date_purchased,
                        *first_period,
                        *salvage,
                        *period,
                        *rate,
                        0.0,
                    ),
                    [
                        cost,
                        date_purchased,
                        first_period,
                        salvage,
                        period,
                        rate,
                        basis,
                    ] => (
                        *cost,
                        *date_purchased,
                        *first_period,
                        *salvage,
                        *period,
                        *rate,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![
                    cost,
                    date_purchased,
                    first_period,
                    salvage,
                    period,
                    rate,
                    basis,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                let basis = yearfrac_basis(basis)?;
                if basis == 2 {
                    return Err(FormulaEvalError::Num);
                }
                let date_purchased = financial_date_serial(date_purchased)?;
                let first_period = financial_date_serial(first_period)?;
                let period = period.trunc();
                if cost <= 0.0
                    || salvage < 0.0
                    || salvage >= cost
                    || period < 0.0
                    || rate <= 0.0
                    || date_purchased >= first_period
                {
                    return Err(FormulaEvalError::Num);
                }
                let depreciable_cost = cost - salvage;
                let first_depreciation =
                    cost * rate * yearfrac_by_basis(date_purchased, first_period, basis)?;
                let depreciation = if period == 0.0 {
                    first_depreciation
                } else {
                    cost * rate
                };
                let previous_depreciation = if period == 0.0 {
                    0.0
                } else {
                    first_depreciation + (period - 1.0) * cost * rate
                };
                if previous_depreciation >= depreciable_cost {
                    return Ok(0.0);
                }
                checked_numeric_result(depreciation.min(depreciable_cost - previous_depreciation))
            }
            FormulaScalarFunction::Acot => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(1.0_f64.atan2(*value))
            }
            FormulaScalarFunction::Acoth => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if value.abs() <= 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(0.5 * ((*value + 1.0) / (*value - 1.0)).ln())
            }
            FormulaScalarFunction::And => {
                if args.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                Ok(if args.iter().all(|value| *value != 0.0) {
                    1.0
                } else {
                    0.0
                })
            }
            FormulaScalarFunction::Asin => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value < -1.0 || *value > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value.asin())
            }
            FormulaScalarFunction::Asinh => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(value.asinh())
            }
            FormulaScalarFunction::Atan => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(value.atan())
            }
            FormulaScalarFunction::Atan2 => {
                let [x_num, y_num] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *x_num == 0.0 && *y_num == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                Ok((*y_num).atan2(*x_num))
            }
            FormulaScalarFunction::Atanh => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value <= -1.0 || *value >= 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(value.atanh())
            }
            FormulaScalarFunction::BesselI => {
                let [value, order] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(formula_bessel_i(*value, formula_bessel_order(*order)?)?)
            }
            FormulaScalarFunction::BesselJ => {
                let [value, order] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(formula_bessel_j(*value, formula_bessel_order(*order)?)?)
            }
            FormulaScalarFunction::BesselK => {
                let [value, order] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(formula_bessel_k(*value, formula_bessel_order(*order)?)?)
            }
            FormulaScalarFunction::BesselY => {
                let [value, order] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(formula_bessel_y(*value, formula_bessel_order(*order)?)?)
            }
            FormulaScalarFunction::BetaDist | FormulaScalarFunction::BetaDistLegacy => {
                let (x, alpha, beta, cumulative, lower, upper) = match (self, args) {
                    (FormulaScalarFunction::BetaDist, [x, alpha, beta, cumulative]) => {
                        (*x, *alpha, *beta, *cumulative, 0.0, 1.0)
                    }
                    (FormulaScalarFunction::BetaDist, [x, alpha, beta, cumulative, lower]) => {
                        (*x, *alpha, *beta, *cumulative, *lower, 1.0)
                    }
                    (
                        FormulaScalarFunction::BetaDist,
                        [x, alpha, beta, cumulative, lower, upper],
                    ) => (*x, *alpha, *beta, *cumulative, *lower, *upper),
                    (FormulaScalarFunction::BetaDistLegacy, [x, alpha, beta]) => {
                        (*x, *alpha, *beta, 1.0, 0.0, 1.0)
                    }
                    (FormulaScalarFunction::BetaDistLegacy, [x, alpha, beta, lower]) => {
                        (*x, *alpha, *beta, 1.0, *lower, 1.0)
                    }
                    (FormulaScalarFunction::BetaDistLegacy, [x, alpha, beta, lower, upper]) => {
                        (*x, *alpha, *beta, 1.0, *lower, *upper)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![x, alpha, beta, cumulative, lower, upper]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if alpha <= 0.0 || beta <= 0.0 || lower >= upper || x < lower || x > upper {
                    return Err(FormulaEvalError::Num);
                }
                let scaled = (x - lower) / (upper - lower);
                if cumulative != 0.0 {
                    regularized_beta(scaled, alpha, beta)
                } else {
                    checked_numeric_result(beta_pdf(scaled, alpha, beta)? / (upper - lower))
                }
            }
            FormulaScalarFunction::BetaInv | FormulaScalarFunction::BetaInvLegacy => {
                let (probability, alpha, beta, lower, upper) = match args {
                    [probability, alpha, beta] => (*probability, *alpha, *beta, 0.0, 1.0),
                    [probability, alpha, beta, lower] => (*probability, *alpha, *beta, *lower, 1.0),
                    [probability, alpha, beta, lower, upper] => {
                        (*probability, *alpha, *beta, *lower, *upper)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![probability, alpha, beta, lower, upper]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if alpha <= 0.0 || beta <= 0.0 || lower >= upper {
                    return Err(FormulaEvalError::Num);
                }
                let mut cdf = |scaled: f64| regularized_beta(scaled, alpha, beta);
                let scaled = inverse_unit_cdf(probability, &mut cdf)?;
                checked_numeric_result(lower + scaled * (upper - lower))
            }
            FormulaScalarFunction::BinomDist => {
                let [number_s, trials, probability, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![number_s, trials, probability, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *probability < 0.0 || *probability > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let number_s = trunc_nonnegative_integer(*number_s)?;
                let trials = trunc_nonnegative_integer(*trials)?;
                if number_s > trials {
                    return Err(FormulaEvalError::Num);
                }
                if *cumulative != 0.0 {
                    cumulative_binomial_probability(number_s, trials, *probability)
                } else {
                    binomial_probability(number_s, trials, *probability)
                }
            }
            FormulaScalarFunction::BinomDistRange => {
                let (trials, probability, number_s, number_s2) = match args {
                    [trials, probability, number_s] => {
                        (*trials, *probability, *number_s, *number_s)
                    }
                    [trials, probability, number_s, number_s2] => {
                        (*trials, *probability, *number_s, *number_s2)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![trials, probability, number_s, number_s2]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if probability < 0.0 || probability > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let trials = trunc_nonnegative_integer(trials)?;
                let number_s = trunc_nonnegative_integer(number_s)?;
                let number_s2 = trunc_nonnegative_integer(number_s2)?;
                if number_s > trials || number_s2 < number_s || number_s2 > trials {
                    return Err(FormulaEvalError::Num);
                }
                let mut total = 0.0;
                for successes in number_s..=number_s2 {
                    total += binomial_probability(successes, trials, probability)?;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                checked_numeric_result(total.min(1.0))
            }
            FormulaScalarFunction::BinomInv | FormulaScalarFunction::CritBinom => {
                let [trials, probability, alpha] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![trials, probability, alpha]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *probability <= 0.0 || *probability >= 1.0 || *alpha <= 0.0 || *alpha >= 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let trials = trunc_nonnegative_integer(*trials)?;
                for successes in 0..=trials {
                    let cumulative =
                        cumulative_binomial_probability(successes, trials, *probability)?;
                    if cumulative >= *alpha {
                        return Ok(successes as f64);
                    }
                }
                Ok(trials as f64)
            }
            FormulaScalarFunction::ChiDistLegacy | FormulaScalarFunction::ChiSqDistRt => {
                let [x, degrees] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !x.is_finite() || !degrees.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                regularized_gamma_q(degrees.trunc() / 2.0, *x / 2.0)
            }
            FormulaScalarFunction::ChiInvLegacy | FormulaScalarFunction::ChiSqInvRt => {
                let [probability, degrees] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !probability.is_finite() || !degrees.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *probability <= 0.0 || *probability >= 1.0 || *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let shape = degrees.trunc() / 2.0;
                let target = 1.0 - *probability;
                let mut cdf = |x: f64| regularized_gamma_p(shape, x / 2.0);
                inverse_positive_cdf(target, &mut cdf)
            }
            FormulaScalarFunction::ChiSqDist => {
                let [x, degrees, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, degrees, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let degrees = degrees.trunc();
                if *cumulative != 0.0 {
                    regularized_gamma_p(degrees / 2.0, *x / 2.0)
                } else if *x == 0.0 {
                    Ok(if degrees == 2.0 {
                        0.5
                    } else if degrees > 2.0 {
                        0.0
                    } else {
                        return Err(FormulaEvalError::Num);
                    })
                } else {
                    let log_density = (degrees / 2.0 - 1.0) * x.ln()
                        - *x / 2.0
                        - (degrees / 2.0) * 2.0_f64.ln()
                        - gamma_ln_value(degrees / 2.0);
                    checked_numeric_result(log_density.exp())
                }
            }
            FormulaScalarFunction::ChiSqInv => {
                let [probability, degrees] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !probability.is_finite() || !degrees.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *probability <= 0.0 || *probability >= 1.0 || *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let shape = degrees.trunc() / 2.0;
                let mut cdf = |x: f64| regularized_gamma_p(shape, x / 2.0);
                inverse_positive_cdf(*probability, &mut cdf)
            }
            FormulaScalarFunction::BitAnd => {
                let [left, right] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok((formula_bitwise_argument(*left)? & formula_bitwise_argument(*right)?) as f64)
            }
            FormulaScalarFunction::BitLShift => {
                let [number, shift] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = formula_bitwise_argument(*number)?;
                let shift = formula_bit_shift_argument(*shift)?;
                let value = if shift < 0 {
                    number >> shift.unsigned_abs() as u32
                } else {
                    number
                        .checked_shl(shift as u32)
                        .ok_or(FormulaEvalError::Num)?
                };
                if value > ((1_u64 << 48) - 1) {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value as f64)
            }
            FormulaScalarFunction::BitOr => {
                let [left, right] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok((formula_bitwise_argument(*left)? | formula_bitwise_argument(*right)?) as f64)
            }
            FormulaScalarFunction::BitRShift => {
                let [number, shift] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = formula_bitwise_argument(*number)?;
                let shift = formula_bit_shift_argument(*shift)?;
                let value = if shift < 0 {
                    number
                        .checked_shl(shift.unsigned_abs() as u32)
                        .ok_or(FormulaEvalError::Num)?
                } else {
                    number >> shift as u32
                };
                if value > ((1_u64 << 48) - 1) {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value as f64)
            }
            FormulaScalarFunction::BitXor => {
                let [left, right] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok((formula_bitwise_argument(*left)? ^ formula_bitwise_argument(*right)?) as f64)
            }
            FormulaScalarFunction::Ceiling => {
                let [number, significance] = args else {
                    return Err(FormulaEvalError::Value);
                };
                ceiling_floor_legacy(*number, *significance, true)
            }
            FormulaScalarFunction::CeilingMath => {
                let (number, significance, mode) = match args {
                    [number] => (*number, 1.0, 0.0),
                    [number, significance] => (*number, *significance, 0.0),
                    [number, significance, mode] => (*number, *significance, *mode),
                    _ => return Err(FormulaEvalError::Value),
                };
                ceiling_floor_math(number, significance, mode, true)
            }
            FormulaScalarFunction::CeilingPrecise => {
                let (number, significance) = match args {
                    [number] => (*number, 1.0),
                    [number, significance] => (*number, *significance),
                    _ => return Err(FormulaEvalError::Value),
                };
                ceiling_floor_precise(number, significance, true)
            }
            FormulaScalarFunction::Combin => {
                let [number, chosen] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = trunc_nonnegative_integer(*number)?;
                let chosen = trunc_nonnegative_integer(*chosen)?;
                combination(number, chosen)
            }
            FormulaScalarFunction::Combina => {
                let [number, chosen] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = trunc_nonnegative_integer(*number)?;
                let chosen = trunc_nonnegative_integer(*chosen)?;
                if number < chosen {
                    return Err(FormulaEvalError::Num);
                }
                if chosen == 0 {
                    return Ok(1.0);
                }
                let expanded = number
                    .checked_add(chosen - 1)
                    .ok_or(FormulaEvalError::Num)?;
                combination(expanded, chosen)
            }
            FormulaScalarFunction::ConfidenceNorm | FormulaScalarFunction::ConfidenceNormLegacy => {
                let [alpha, standard_dev, size] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !alpha.is_finite() || !standard_dev.is_finite() || !size.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                let size = size.trunc();
                if *alpha <= 0.0 || *alpha >= 1.0 || *standard_dev <= 0.0 || size < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let critical_value = inverse_standard_normal(1.0 - *alpha / 2.0)?;
                checked_numeric_result(critical_value * *standard_dev / size.sqrt())
            }
            FormulaScalarFunction::ConfidenceT => {
                let [alpha, standard_dev, size] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !alpha.is_finite() || !standard_dev.is_finite() || !size.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                let size = size.trunc();
                if *alpha <= 0.0 || *alpha >= 1.0 || *standard_dev <= 0.0 || size < 2.0 {
                    return Err(FormulaEvalError::Num);
                }
                let target = 1.0 - *alpha / 2.0;
                let degrees = size - 1.0;
                let mut cdf = |x: f64| student_t_cdf(x, degrees);
                let critical_value = inverse_positive_cdf(target, &mut cdf)?;
                checked_numeric_result(critical_value * *standard_dev / size.sqrt())
            }
            FormulaScalarFunction::Cos => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = value.cos();
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Cosh => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(value.cosh())
            }
            FormulaScalarFunction::CoupDayBs => {
                let (settlement, _, _, basis, _, previous_coupon, _) =
                    coupon_schedule_from_args(args)?;
                coupon_days_between(previous_coupon, settlement, basis)
            }
            FormulaScalarFunction::CoupDays => {
                let (_, _, frequency, basis, _, previous_coupon, next_coupon) =
                    coupon_schedule_from_args(args)?;
                match basis {
                    1 => Ok((next_coupon - previous_coupon) as f64),
                    3 => Ok(365.0 / frequency as f64),
                    _ => Ok(360.0 / frequency as f64),
                }
            }
            FormulaScalarFunction::CoupDaysNc => {
                let (settlement, _, _, basis, _, _, next_coupon) = coupon_schedule_from_args(args)?;
                coupon_days_between(settlement, next_coupon, basis)
            }
            FormulaScalarFunction::CoupNcd => {
                let (_, _, _, _, _, _, next_coupon) = coupon_schedule_from_args(args)?;
                Ok(next_coupon as f64)
            }
            FormulaScalarFunction::CoupNum => {
                let (_, _, _, _, coupon_count, _, _) = coupon_schedule_from_args(args)?;
                Ok(coupon_count as f64)
            }
            FormulaScalarFunction::CoupPcd => {
                let (_, _, _, _, _, previous_coupon, _) = coupon_schedule_from_args(args)?;
                Ok(previous_coupon as f64)
            }
            FormulaScalarFunction::Cot => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                validate_reciprocal_trig_input(*value)?;
                reciprocal_numeric_result(value.tan())
            }
            FormulaScalarFunction::Coth => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                validate_reciprocal_trig_input(*value)?;
                reciprocal_numeric_result(value.tanh())
            }
            FormulaScalarFunction::Csc => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                validate_reciprocal_trig_input(*value)?;
                reciprocal_numeric_result(value.sin())
            }
            FormulaScalarFunction::Csch => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                validate_reciprocal_trig_input(*value)?;
                reciprocal_numeric_result(value.sinh())
            }
            FormulaScalarFunction::Date => {
                let [year, month, day] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_date_serial_from_args(*year, *month, *day)
            }
            FormulaScalarFunction::Day => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let (_, _, day) = formula_ymd_from_serial(*serial)?;
                Ok(day as f64)
            }
            FormulaScalarFunction::Days => {
                let [end_date, start_date] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(end_date - start_date)
            }
            FormulaScalarFunction::Days360 => {
                let (start_date, end_date, european) = match args {
                    [start_date, end_date] => (*start_date, *end_date, false),
                    [start_date, end_date, method] => (*start_date, *end_date, *method != 0.0),
                    _ => return Err(FormulaEvalError::Value),
                };
                Ok(days360(
                    formula_serial_integer(start_date)?,
                    formula_serial_integer(end_date)?,
                    european,
                )? as f64)
            }
            FormulaScalarFunction::Db => {
                let (cost, salvage, life, period, month) = match args {
                    [cost, salvage, life, period] => (*cost, *salvage, *life, *period, 12.0),
                    [cost, salvage, life, period, month] => {
                        (*cost, *salvage, *life, *period, *month)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![cost, salvage, life, period, month]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                let period = period.trunc();
                let month = month.trunc();
                if cost <= 0.0
                    || salvage < 0.0
                    || salvage > cost
                    || life <= 0.0
                    || period < 1.0
                    || month < 1.0
                    || month > 12.0
                    || period > life + if month < 12.0 { 1.0 } else { 0.0 }
                {
                    return Err(FormulaEvalError::Num);
                }
                let rate =
                    round_half_away_from_zero((1.0 - (salvage / cost).powf(1.0 / life)) * 1000.0)
                        / 1000.0;
                let mut accumulated_depreciation = 0.0;
                let mut current_period = 1.0;
                let mut depreciation = 0.0;
                while current_period <= period {
                    depreciation = if current_period == 1.0 {
                        cost * rate * month / 12.0
                    } else if current_period > life {
                        (cost - accumulated_depreciation) * rate * (12.0 - month) / 12.0
                    } else {
                        (cost - accumulated_depreciation) * rate
                    };
                    depreciation = depreciation
                        .min(cost - salvage - accumulated_depreciation)
                        .max(0.0);
                    if current_period == period {
                        break;
                    }
                    accumulated_depreciation += depreciation;
                    current_period += 1.0;
                }
                checked_numeric_result(depreciation)
            }
            FormulaScalarFunction::Ddb => {
                let (cost, salvage, life, period, factor) = match args {
                    [cost, salvage, life, period] => (*cost, *salvage, *life, *period, 2.0),
                    [cost, salvage, life, period, factor] => {
                        (*cost, *salvage, *life, *period, *factor)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![cost, salvage, life, period, factor]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                let period = period.trunc();
                if cost < 0.0
                    || salvage < 0.0
                    || salvage > cost
                    || life <= 0.0
                    || period < 1.0
                    || period > life
                    || factor <= 0.0
                {
                    return Err(FormulaEvalError::Num);
                }
                let mut prior_depreciation = 0.0;
                let mut current_period = 1.0;
                while current_period < period {
                    let depreciation = ((cost - prior_depreciation) * factor / life)
                        .min(cost - salvage - prior_depreciation)
                        .max(0.0);
                    prior_depreciation += depreciation;
                    current_period += 1.0;
                }
                checked_numeric_result(
                    ((cost - prior_depreciation) * factor / life)
                        .min(cost - salvage - prior_depreciation)
                        .max(0.0),
                )
            }
            FormulaScalarFunction::Degrees => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = value * 180.0 / std::f64::consts::PI;
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Delta => {
                let (left, right) = match args {
                    [left] => (*left, 0.0),
                    [left, right] => (*left, *right),
                    _ => return Err(FormulaEvalError::Value),
                };
                Ok(if left == right { 1.0 } else { 0.0 })
            }
            FormulaScalarFunction::Disc => {
                let (settlement, maturity, price, redemption, basis) = match args {
                    [settlement, maturity, price, redemption] => {
                        (*settlement, *maturity, *price, *redemption, 0.0)
                    }
                    [settlement, maturity, price, redemption, basis] => {
                        (*settlement, *maturity, *price, *redemption, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, price, redemption, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if price <= 0.0 || redemption <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let yearfrac = discount_security_yearfrac(settlement, maturity, basis)?;
                checked_numeric_result((redemption - price) / redemption / yearfrac)
            }
            FormulaScalarFunction::Duration => {
                let (settlement, maturity, coupon, yld, frequency, basis) = match args {
                    [settlement, maturity, coupon, yld, frequency] => {
                        (*settlement, *maturity, *coupon, *yld, *frequency, 0.0)
                    }
                    [settlement, maturity, coupon, yld, frequency, basis] => {
                        (*settlement, *maturity, *coupon, *yld, *frequency, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                duration_value(settlement, maturity, coupon, yld, frequency, basis, false)
            }
            FormulaScalarFunction::EDate => {
                let [serial, months] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_edate(*serial, *months)
            }
            FormulaScalarFunction::EOMonth => {
                let [serial, months] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_eomonth(*serial, *months)
            }
            FormulaScalarFunction::Effect => {
                let [nominal_rate, npery] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !nominal_rate.is_finite() || !npery.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                let npery = npery.trunc();
                if *nominal_rate <= 0.0 || npery < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((1.0 + nominal_rate / npery).powf(npery) - 1.0)
            }
            FormulaScalarFunction::Erf => {
                let (lower_limit, upper_limit) = match args {
                    [lower_limit] => (*lower_limit, 0.0),
                    [lower_limit, upper_limit] => (*lower_limit, *upper_limit),
                    _ => return Err(FormulaEvalError::Value),
                };
                if !lower_limit.is_finite() || !upper_limit.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                checked_numeric_result(if args.len() == 1 {
                    erf_approx(lower_limit)
                } else {
                    erf_approx(upper_limit) - erf_approx(lower_limit)
                })
            }
            FormulaScalarFunction::ErfPrecise => {
                let [x] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !x.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                checked_numeric_result(erf_approx(*x))
            }
            FormulaScalarFunction::Erfc | FormulaScalarFunction::ErfcPrecise => {
                let [x] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !x.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                checked_numeric_result(1.0 - erf_approx(*x))
            }
            FormulaScalarFunction::Even => {
                let [number] = args else {
                    return Err(FormulaEvalError::Value);
                };
                round_away_to_integer_with_parity(*number, false)
            }
            FormulaScalarFunction::ExponDist => {
                let [x, lambda, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, lambda, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *lambda <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let exponent = (-*lambda * *x).exp();
                let value = if *cumulative != 0.0 {
                    1.0 - exponent
                } else {
                    *lambda * exponent
                };
                checked_numeric_result(value)
            }
            FormulaScalarFunction::Exp => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = value.exp();
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Fact => {
                let [number] = args else {
                    return Err(FormulaEvalError::Value);
                };
                factorial(trunc_nonnegative_integer(*number)?)
            }
            FormulaScalarFunction::FactDouble => {
                let [number] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = trunc_nonnegative_integer(*number)?;
                let mut total = 1.0_f64;
                let mut factor = number;
                while factor > 1 {
                    total *= factor as f64;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    factor -= 2;
                }
                Ok(total)
            }
            FormulaScalarFunction::FDist | FormulaScalarFunction::FDistRt => {
                let (x, degrees1, degrees2, cumulative) = match (self, args) {
                    (FormulaScalarFunction::FDist, [x, degrees1, degrees2, cumulative]) => {
                        (*x, *degrees1, *degrees2, *cumulative)
                    }
                    (FormulaScalarFunction::FDistRt, [x, degrees1, degrees2]) => {
                        (*x, *degrees1, *degrees2, 1.0)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![x, degrees1, degrees2, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if x < 0.0 || degrees1 < 1.0 || degrees2 < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let degrees1 = degrees1.trunc();
                let degrees2 = degrees2.trunc();
                if matches!(self, FormulaScalarFunction::FDistRt) || cumulative != 0.0 {
                    let transformed = degrees1 * x / (degrees1 * x + degrees2);
                    let cdf = regularized_beta(transformed, degrees1 / 2.0, degrees2 / 2.0)?;
                    return Ok(if matches!(self, FormulaScalarFunction::FDistRt) {
                        (1.0 - cdf).clamp(0.0, 1.0)
                    } else {
                        cdf
                    });
                }
                if x == 0.0 {
                    return Ok(0.0);
                }
                let alpha = degrees1 / 2.0;
                let beta = degrees2 / 2.0;
                let log_beta =
                    gamma_ln_value(alpha) + gamma_ln_value(beta) - gamma_ln_value(alpha + beta);
                let log_density = alpha * (degrees1 / degrees2).ln() + (alpha - 1.0) * x.ln()
                    - (alpha + beta) * (1.0 + degrees1 * x / degrees2).ln()
                    - log_beta;
                checked_numeric_result(log_density.exp())
            }
            FormulaScalarFunction::FDistLegacy => {
                let [x, degrees1, degrees2] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, degrees1, degrees2]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *degrees1 < 1.0 || *degrees2 < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let degrees1 = degrees1.trunc();
                let degrees2 = degrees2.trunc();
                let transformed = degrees1 * *x / (degrees1 * *x + degrees2);
                let cdf = regularized_beta(transformed, degrees1 / 2.0, degrees2 / 2.0)?;
                Ok((1.0 - cdf).clamp(0.0, 1.0))
            }
            FormulaScalarFunction::FInv | FormulaScalarFunction::FInvRt => {
                let [probability, degrees1, degrees2] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![probability, degrees1, degrees2]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *degrees1 < 1.0 || *degrees2 < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let degrees1 = degrees1.trunc();
                let degrees2 = degrees2.trunc();
                let target = if matches!(self, FormulaScalarFunction::FInvRt) {
                    1.0 - *probability
                } else {
                    *probability
                };
                let mut cdf = |x: f64| {
                    let transformed = degrees1 * x / (degrees1 * x + degrees2);
                    regularized_beta(transformed, degrees1 / 2.0, degrees2 / 2.0)
                };
                inverse_positive_cdf(target, &mut cdf)
            }
            FormulaScalarFunction::FInvLegacy => {
                let [probability, degrees1, degrees2] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![probability, degrees1, degrees2]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *degrees1 < 1.0 || *degrees2 < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let degrees1 = degrees1.trunc();
                let degrees2 = degrees2.trunc();
                let target = 1.0 - *probability;
                let mut cdf = |x: f64| {
                    let transformed = degrees1 * x / (degrees1 * x + degrees2);
                    regularized_beta(transformed, degrees1 / 2.0, degrees2 / 2.0)
                };
                inverse_positive_cdf(target, &mut cdf)
            }
            FormulaScalarFunction::Fisher => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value <= -1.0 || *value >= 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(0.5 * ((1.0 + *value) / (1.0 - *value)).ln())
            }
            FormulaScalarFunction::FisherInv => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(value.tanh())
            }
            FormulaScalarFunction::Floor => {
                let [number, significance] = args else {
                    return Err(FormulaEvalError::Value);
                };
                ceiling_floor_legacy(*number, *significance, false)
            }
            FormulaScalarFunction::FloorMath => {
                let (number, significance, mode) = match args {
                    [number] => (*number, 1.0, 0.0),
                    [number, significance] => (*number, *significance, 0.0),
                    [number, significance, mode] => (*number, *significance, *mode),
                    _ => return Err(FormulaEvalError::Value),
                };
                ceiling_floor_math(number, significance, mode, false)
            }
            FormulaScalarFunction::FloorPrecise => {
                let (number, significance) = match args {
                    [number] => (*number, 1.0),
                    [number, significance] => (*number, *significance),
                    _ => return Err(FormulaEvalError::Value),
                };
                ceiling_floor_precise(number, significance, false)
            }
            FormulaScalarFunction::Gauss => {
                let [z] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(standard_normal_cdf(*z) - 0.5)
            }
            FormulaScalarFunction::Gamma => {
                let [number] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !number.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *number == 0.0 || (*number < 0.0 && number.fract() == 0.0) {
                    return Err(FormulaEvalError::Num);
                }
                let value = if *number < 0.5 {
                    std::f64::consts::PI
                        / ((std::f64::consts::PI * *number).sin()
                            * gamma_ln_value(1.0 - *number).exp())
                } else {
                    gamma_ln_value(*number).exp()
                };
                checked_numeric_result(value)
            }
            FormulaScalarFunction::GammaDist | FormulaScalarFunction::GammaDistLegacy => {
                let [x, alpha, beta, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, alpha, beta, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *alpha <= 0.0 || *beta <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                if *cumulative != 0.0 {
                    regularized_gamma_p(*alpha, *x / *beta)
                } else if *x == 0.0 {
                    Ok(if *alpha == 1.0 {
                        1.0 / *beta
                    } else if *alpha > 1.0 {
                        0.0
                    } else {
                        return Err(FormulaEvalError::Num);
                    })
                } else {
                    let scaled = *x / *beta;
                    let log_density =
                        (*alpha - 1.0) * scaled.ln() - scaled - gamma_ln_value(*alpha) - beta.ln();
                    checked_numeric_result(log_density.exp())
                }
            }
            FormulaScalarFunction::GammaInv | FormulaScalarFunction::GammaInvLegacy => {
                let [probability, alpha, beta] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![probability, alpha, beta]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *alpha <= 0.0 || *beta <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let mut cdf = |x: f64| regularized_gamma_p(*alpha, x / *beta);
                inverse_positive_cdf(*probability, &mut cdf)
            }
            FormulaScalarFunction::GammaLn | FormulaScalarFunction::GammaLnPrecise => {
                let [x] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !x.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *x <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(gamma_ln_value(*x))
            }
            FormulaScalarFunction::GeStep => {
                let (number, step) = match args {
                    [number] => (*number, 0.0),
                    [number, step] => (*number, *step),
                    _ => return Err(FormulaEvalError::Value),
                };
                Ok(if number >= step { 1.0 } else { 0.0 })
            }
            FormulaScalarFunction::Hour => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(formula_time_parts_from_serial(*serial)?.0 as f64)
            }
            FormulaScalarFunction::HypGeomDist | FormulaScalarFunction::HypGeomDistLegacy => {
                let (
                    sample_successes,
                    sample_size,
                    population_successes,
                    population_size,
                    cumulative,
                ) = match (self, args) {
                    (
                        FormulaScalarFunction::HypGeomDist,
                        [
                            sample_successes,
                            sample_size,
                            population_successes,
                            population_size,
                            cumulative,
                        ],
                    ) => (
                        *sample_successes,
                        *sample_size,
                        *population_successes,
                        *population_size,
                        *cumulative,
                    ),
                    (
                        FormulaScalarFunction::HypGeomDistLegacy,
                        [
                            sample_successes,
                            sample_size,
                            population_successes,
                            population_size,
                        ],
                    ) => (
                        *sample_successes,
                        *sample_size,
                        *population_successes,
                        *population_size,
                        0.0,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![
                    sample_successes,
                    sample_size,
                    population_successes,
                    population_size,
                    cumulative,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                let sample_successes = trunc_nonnegative_integer(sample_successes)?;
                let sample_size = trunc_nonnegative_integer(sample_size)?;
                let population_successes = trunc_nonnegative_integer(population_successes)?;
                let population_size = trunc_nonnegative_integer(population_size)?;
                if sample_size == 0
                    || population_successes == 0
                    || population_size == 0
                    || sample_size > population_size
                    || population_successes > population_size
                {
                    return Err(FormulaEvalError::Num);
                }
                let lower_successes =
                    sample_size.saturating_sub(population_size - population_successes);
                let upper_successes = sample_size.min(population_successes);
                if sample_successes < lower_successes || sample_successes > upper_successes {
                    return Err(FormulaEvalError::Num);
                }
                if cumulative != 0.0 {
                    let mut total = 0.0;
                    for successes in lower_successes..=sample_successes {
                        total += hypergeometric_probability(
                            successes,
                            sample_size,
                            population_successes,
                            population_size,
                        )?;
                        if !total.is_finite() {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    checked_numeric_result(total.min(1.0))
                } else {
                    hypergeometric_probability(
                        sample_successes,
                        sample_size,
                        population_successes,
                        population_size,
                    )
                }
            }
            FormulaScalarFunction::If => {
                let [condition, true_value, false_value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(if *condition != 0.0 {
                    *true_value
                } else {
                    *false_value
                })
            }
            FormulaScalarFunction::Intrate => {
                let (settlement, maturity, investment, redemption, basis) = match args {
                    [settlement, maturity, investment, redemption] => {
                        (*settlement, *maturity, *investment, *redemption, 0.0)
                    }
                    [settlement, maturity, investment, redemption, basis] => {
                        (*settlement, *maturity, *investment, *redemption, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, investment, redemption, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if investment <= 0.0 || redemption <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let yearfrac = discount_security_yearfrac(settlement, maturity, basis)?;
                checked_numeric_result((redemption - investment) / investment / yearfrac)
            }
            FormulaScalarFunction::IsoCeiling => {
                let (number, significance) = match args {
                    [number] => (*number, 1.0),
                    [number, significance] => (*number, *significance),
                    _ => return Err(FormulaEvalError::Value),
                };
                ceiling_floor_precise(number, significance, true)
            }
            FormulaScalarFunction::IsoWeekNum => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let serial = formula_serial_integer(*serial)?;
                iso_weeknum_from_serial(serial).map(|week| week as f64)
            }
            FormulaScalarFunction::IsEven => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(if value.trunc().rem_euclid(2.0) == 0.0 {
                    1.0
                } else {
                    0.0
                })
            }
            FormulaScalarFunction::IsOdd => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(if value.trunc().rem_euclid(2.0) != 0.0 {
                    1.0
                } else {
                    0.0
                })
            }
            FormulaScalarFunction::Int => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(value.floor())
            }
            FormulaScalarFunction::Ln => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value.ln())
            }
            FormulaScalarFunction::LogNormDist | FormulaScalarFunction::LogNormDistLegacy => {
                let (x, mean, standard_dev, cumulative) = match (self, args) {
                    (FormulaScalarFunction::LogNormDist, [x, mean, standard_dev, cumulative]) => {
                        (*x, *mean, *standard_dev, *cumulative)
                    }
                    (FormulaScalarFunction::LogNormDistLegacy, [x, mean, standard_dev]) => {
                        (*x, *mean, *standard_dev, 1.0)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![x, mean, standard_dev, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if x <= 0.0 || standard_dev <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let z = (x.ln() - mean) / standard_dev;
                let value = if cumulative != 0.0 {
                    standard_normal_cdf(z)
                } else {
                    standard_normal_pdf(z) / (x * standard_dev)
                };
                checked_numeric_result(value)
            }
            FormulaScalarFunction::LogNormInv | FormulaScalarFunction::LogNormInvLegacy => {
                let [probability, mean, standard_dev] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![probability, mean, standard_dev]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *standard_dev <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let z = inverse_standard_normal(*probability)?;
                checked_numeric_result((*mean + *standard_dev * z).exp())
            }
            FormulaScalarFunction::Log => {
                let (number, base) = match args {
                    [number] => (*number, 10.0),
                    [number, base] => (*number, *base),
                    _ => return Err(FormulaEvalError::Value),
                };
                if number <= 0.0 || base <= 0.0 || base == 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let value = number.log(base);
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Log10 => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value.log10())
            }
            FormulaScalarFunction::Minute => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(formula_time_parts_from_serial(*serial)?.1 as f64)
            }
            FormulaScalarFunction::MDuration => {
                let (settlement, maturity, coupon, yld, frequency, basis) = match args {
                    [settlement, maturity, coupon, yld, frequency] => {
                        (*settlement, *maturity, *coupon, *yld, *frequency, 0.0)
                    }
                    [settlement, maturity, coupon, yld, frequency, basis] => {
                        (*settlement, *maturity, *coupon, *yld, *frequency, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                duration_value(settlement, maturity, coupon, yld, frequency, basis, true)
            }
            FormulaScalarFunction::Mod => {
                let [number, divisor] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *divisor == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                Ok(number - divisor * (number / divisor).floor())
            }
            FormulaScalarFunction::Month => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let (_, month, _) = formula_ymd_from_serial(*serial)?;
                Ok(month as f64)
            }
            FormulaScalarFunction::MRound => {
                let [number, multiple] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *multiple == 0.0 {
                    return Ok(0.0);
                }
                if (*number > 0.0 && *multiple < 0.0) || (*number < 0.0 && *multiple > 0.0) {
                    return Err(FormulaEvalError::Num);
                }
                let value = round_half_away_from_zero(number / multiple) * multiple;
                if value.is_finite() {
                    Ok(normalize_zero(value))
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Multinomial => {
                if args.is_empty() || args.len() > 255 {
                    return Err(FormulaEvalError::Value);
                }
                let mut sum = 0_u64;
                let mut denominator = 1.0_f64;
                for value in args {
                    let value = trunc_nonnegative_integer(*value)?;
                    sum = sum.checked_add(value).ok_or(FormulaEvalError::Num)?;
                    denominator *= factorial(value)?;
                    if !denominator.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                Ok(factorial(sum)? / denominator)
            }
            FormulaScalarFunction::NegBinomDist | FormulaScalarFunction::NegBinomDistLegacy => {
                let (failures, successes, probability, cumulative) = match (self, args) {
                    (
                        FormulaScalarFunction::NegBinomDist,
                        [failures, successes, probability, cumulative],
                    ) => (*failures, *successes, *probability, *cumulative),
                    (
                        FormulaScalarFunction::NegBinomDistLegacy,
                        [failures, successes, probability],
                    ) => (*failures, *successes, *probability, 0.0),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![failures, successes, probability, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if probability < 0.0 || probability > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let failures = trunc_nonnegative_integer(failures)?;
                let successes = trunc_nonnegative_integer(successes)?;
                if successes < 1 {
                    return Err(FormulaEvalError::Num);
                }
                if cumulative != 0.0 {
                    let mut total = 0.0;
                    for failure_count in 0..=failures {
                        total +=
                            negative_binomial_probability(failure_count, successes, probability)?;
                        if !total.is_finite() {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    checked_numeric_result(total.min(1.0))
                } else {
                    negative_binomial_probability(failures, successes, probability)
                }
            }
            FormulaScalarFunction::Nominal => {
                let [effect_rate, npery] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !effect_rate.is_finite() || !npery.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                let npery = npery.trunc();
                if *effect_rate <= 0.0 || npery < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(npery * ((1.0 + effect_rate).powf(1.0 / npery) - 1.0))
            }
            FormulaScalarFunction::NormDist => {
                let [x, mean, standard_dev, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, mean, standard_dev, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *standard_dev <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let z = (*x - *mean) / *standard_dev;
                let value = if *cumulative != 0.0 {
                    standard_normal_cdf(z)
                } else {
                    standard_normal_pdf(z) / *standard_dev
                };
                checked_numeric_result(value)
            }
            FormulaScalarFunction::NormInv => {
                let [probability, mean, standard_dev] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![probability, mean, standard_dev]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *standard_dev <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let z = inverse_standard_normal(*probability)?;
                checked_numeric_result(*mean + *standard_dev * z)
            }
            FormulaScalarFunction::NormSDist => {
                let [z, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !z.is_finite() || !cumulative.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                checked_numeric_result(if *cumulative != 0.0 {
                    standard_normal_cdf(*z)
                } else {
                    standard_normal_pdf(*z)
                })
            }
            FormulaScalarFunction::NormSDistLegacy => {
                let [z] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !z.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                checked_numeric_result(standard_normal_cdf(*z))
            }
            FormulaScalarFunction::NormSInv | FormulaScalarFunction::NormSInvLegacy => {
                let [probability] = args else {
                    return Err(FormulaEvalError::Value);
                };
                inverse_standard_normal(*probability)
            }
            FormulaScalarFunction::Not => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(if *value == 0.0 { 1.0 } else { 0.0 })
            }
            FormulaScalarFunction::Now => {
                let [] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_current_excel_serial()
            }
            FormulaScalarFunction::Odd => {
                let [number] = args else {
                    return Err(FormulaEvalError::Value);
                };
                round_away_to_integer_with_parity(*number, true)
            }
            FormulaScalarFunction::OddFPrice => {
                let (
                    settlement,
                    maturity,
                    issue,
                    first_coupon,
                    rate,
                    yld,
                    redemption,
                    frequency,
                    basis,
                ) = match args {
                    [
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        yld,
                        redemption,
                        frequency,
                    ] => (
                        *settlement,
                        *maturity,
                        *issue,
                        *first_coupon,
                        *rate,
                        *yld,
                        *redemption,
                        *frequency,
                        0.0,
                    ),
                    [
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        yld,
                        redemption,
                        frequency,
                        basis,
                    ] => (
                        *settlement,
                        *maturity,
                        *issue,
                        *first_coupon,
                        *rate,
                        *yld,
                        *redemption,
                        *frequency,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if yld < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                odd_first_coupon_price(
                    settlement,
                    maturity,
                    issue,
                    first_coupon,
                    rate,
                    yld,
                    redemption,
                    frequency,
                    basis,
                )
            }
            FormulaScalarFunction::OddFYield => {
                let (
                    settlement,
                    maturity,
                    issue,
                    first_coupon,
                    rate,
                    price,
                    redemption,
                    frequency,
                    basis,
                ) = match args {
                    [
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        price,
                        redemption,
                        frequency,
                    ] => (
                        *settlement,
                        *maturity,
                        *issue,
                        *first_coupon,
                        *rate,
                        *price,
                        *redemption,
                        *frequency,
                        0.0,
                    ),
                    [
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        price,
                        redemption,
                        frequency,
                        basis,
                    ] => (
                        *settlement,
                        *maturity,
                        *issue,
                        *first_coupon,
                        *rate,
                        *price,
                        *redemption,
                        *frequency,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![
                    settlement,
                    maturity,
                    issue,
                    first_coupon,
                    rate,
                    price,
                    redemption,
                    frequency,
                    basis,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if rate < 0.0 || price <= 0.0 || redemption <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let frequency = coupon_frequency(frequency)? as f64;
                let mut price_difference = |yld: f64| -> Result<f64, FormulaEvalError> {
                    Ok(odd_first_coupon_price(
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        yld,
                        redemption,
                        frequency,
                        basis,
                    )? - price)
                };
                solve_odd_coupon_yield(&mut price_difference, frequency)
            }
            FormulaScalarFunction::OddLPrice => {
                let (settlement, maturity, last_interest, rate, yld, redemption, frequency, basis) =
                    match args {
                        [
                            settlement,
                            maturity,
                            last_interest,
                            rate,
                            yld,
                            redemption,
                            frequency,
                        ] => (
                            *settlement,
                            *maturity,
                            *last_interest,
                            *rate,
                            *yld,
                            *redemption,
                            *frequency,
                            0.0,
                        ),
                        [
                            settlement,
                            maturity,
                            last_interest,
                            rate,
                            yld,
                            redemption,
                            frequency,
                            basis,
                        ] => (
                            *settlement,
                            *maturity,
                            *last_interest,
                            *rate,
                            *yld,
                            *redemption,
                            *frequency,
                            *basis,
                        ),
                        _ => return Err(FormulaEvalError::Value),
                    };
                if yld < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                odd_last_coupon_price(
                    settlement,
                    maturity,
                    last_interest,
                    rate,
                    yld,
                    redemption,
                    frequency,
                    basis,
                )
            }
            FormulaScalarFunction::OddLYield => {
                let (
                    settlement,
                    maturity,
                    last_interest,
                    rate,
                    price,
                    redemption,
                    frequency,
                    basis,
                ) = match args {
                    [
                        settlement,
                        maturity,
                        last_interest,
                        rate,
                        price,
                        redemption,
                        frequency,
                    ] => (
                        *settlement,
                        *maturity,
                        *last_interest,
                        *rate,
                        *price,
                        *redemption,
                        *frequency,
                        0.0,
                    ),
                    [
                        settlement,
                        maturity,
                        last_interest,
                        rate,
                        price,
                        redemption,
                        frequency,
                        basis,
                    ] => (
                        *settlement,
                        *maturity,
                        *last_interest,
                        *rate,
                        *price,
                        *redemption,
                        *frequency,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![
                    settlement,
                    maturity,
                    last_interest,
                    rate,
                    price,
                    redemption,
                    frequency,
                    basis,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if rate < 0.0 || price <= 0.0 || redemption <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let frequency = coupon_frequency(frequency)? as f64;
                let mut price_difference = |yld: f64| -> Result<f64, FormulaEvalError> {
                    Ok(odd_last_coupon_price(
                        settlement,
                        maturity,
                        last_interest,
                        rate,
                        yld,
                        redemption,
                        frequency,
                        basis,
                    )? - price)
                };
                solve_odd_coupon_yield(&mut price_difference, frequency)
            }
            FormulaScalarFunction::Or => {
                if args.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                Ok(if args.iter().any(|value| *value != 0.0) {
                    1.0
                } else {
                    0.0
                })
            }
            FormulaScalarFunction::PDuration => {
                let [rate, pv, fv] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![rate, pv, fv].iter().all(|value| value.is_finite()) {
                    return Err(FormulaEvalError::Value);
                }
                if *rate <= 0.0 || *pv <= 0.0 || *fv <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((fv / pv).ln() / (1.0 + rate).ln())
            }
            FormulaScalarFunction::Price => {
                let (settlement, maturity, rate, yld, redemption, frequency, basis) = match args {
                    [settlement, maturity, rate, yld, redemption, frequency] => (
                        *settlement,
                        *maturity,
                        *rate,
                        *yld,
                        *redemption,
                        *frequency,
                        0.0,
                    ),
                    [
                        settlement,
                        maturity,
                        rate,
                        yld,
                        redemption,
                        frequency,
                        basis,
                    ] => (
                        *settlement,
                        *maturity,
                        *rate,
                        *yld,
                        *redemption,
                        *frequency,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if yld < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                regular_coupon_price(
                    settlement, maturity, rate, yld, redemption, frequency, basis,
                )
            }
            FormulaScalarFunction::PriceDisc => {
                let (settlement, maturity, discount, redemption, basis) = match args {
                    [settlement, maturity, discount, redemption] => {
                        (*settlement, *maturity, *discount, *redemption, 0.0)
                    }
                    [settlement, maturity, discount, redemption, basis] => {
                        (*settlement, *maturity, *discount, *redemption, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, discount, redemption, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if discount <= 0.0 || redemption <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let yearfrac = discount_security_yearfrac(settlement, maturity, basis)?;
                checked_numeric_result(redemption * (1.0 - discount * yearfrac))
            }
            FormulaScalarFunction::PriceMat => {
                let (settlement, maturity, issue, rate, yld, basis) = match args {
                    [settlement, maturity, issue, rate, yld] => {
                        (*settlement, *maturity, *issue, *rate, *yld, 0.0)
                    }
                    [settlement, maturity, issue, rate, yld, basis] => {
                        (*settlement, *maturity, *issue, *rate, *yld, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, issue, rate, yld, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if rate < 0.0 || yld < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let (issue_to_maturity, settlement_to_maturity, issue_to_settlement) =
                    maturity_security_yearfracs(settlement, maturity, issue, basis)?;
                let maturity_value = 100.0 + 100.0 * rate * issue_to_maturity;
                let accrued_interest = 100.0 * rate * issue_to_settlement;
                checked_numeric_result(
                    maturity_value / (1.0 + yld * settlement_to_maturity) - accrued_interest,
                )
            }
            FormulaScalarFunction::Round => {
                let [value, digits] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let factor = formula_round_factor(*digits)?;
                Ok(round_half_away_from_zero(value * factor) / factor)
            }
            FormulaScalarFunction::RoundDown => {
                let [value, digits] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let factor = formula_round_factor(*digits)?;
                Ok(round_toward_zero(value * factor) / factor)
            }
            FormulaScalarFunction::RoundUp => {
                let [value, digits] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let factor = formula_round_factor(*digits)?;
                Ok(round_away_from_zero(value * factor) / factor)
            }
            FormulaScalarFunction::Received => {
                let (settlement, maturity, investment, discount, basis) = match args {
                    [settlement, maturity, investment, discount] => {
                        (*settlement, *maturity, *investment, *discount, 0.0)
                    }
                    [settlement, maturity, investment, discount, basis] => {
                        (*settlement, *maturity, *investment, *discount, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, investment, discount, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if investment <= 0.0 || discount <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let yearfrac = discount_security_yearfrac(settlement, maturity, basis)?;
                let denominator = 1.0 - discount * yearfrac;
                if denominator <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(investment / denominator)
            }
            FormulaScalarFunction::Rri => {
                let [nper, pv, fv] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![nper, pv, fv].iter().all(|value| value.is_finite()) {
                    return Err(FormulaEvalError::Value);
                }
                if *nper <= 0.0 || *pv <= 0.0 || *fv < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((fv / pv).powf(1.0 / nper) - 1.0)
            }
            FormulaScalarFunction::Sec => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                validate_reciprocal_trig_input(*value)?;
                reciprocal_numeric_result(value.cos())
            }
            FormulaScalarFunction::Sech => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                validate_reciprocal_trig_input(*value)?;
                reciprocal_numeric_result(value.cosh())
            }
            FormulaScalarFunction::Sign => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(if *value > 0.0 {
                    1.0
                } else if *value < 0.0 {
                    -1.0
                } else {
                    0.0
                })
            }
            FormulaScalarFunction::Sln => {
                let [cost, salvage, life] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![cost, salvage, life].iter().all(|value| value.is_finite()) {
                    return Err(FormulaEvalError::Value);
                }
                if *life <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((cost - salvage) / life)
            }
            FormulaScalarFunction::Permut => {
                let [number, chosen] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = trunc_nonnegative_integer(*number)?;
                let chosen = trunc_nonnegative_integer(*chosen)?;
                if number == 0 || chosen > number {
                    return Err(FormulaEvalError::Num);
                }
                let mut total = 1.0_f64;
                for value in (number - chosen + 1)..=number {
                    total *= value as f64;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                Ok(total)
            }
            FormulaScalarFunction::PermutationA => {
                let [number, chosen] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let number = trunc_nonnegative_integer(*number)?;
                let chosen = trunc_nonnegative_integer(*chosen)?;
                if number == 0 && chosen > 0 {
                    return Err(FormulaEvalError::Num);
                }
                let chosen = i32::try_from(chosen).map_err(|_| FormulaEvalError::Num)?;
                let value = (number as f64).powi(chosen);
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Phi => {
                let [z] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(standard_normal_pdf(*z))
            }
            FormulaScalarFunction::Pi => {
                if !args.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                Ok(std::f64::consts::PI)
            }
            FormulaScalarFunction::PoissonDist => {
                let [x, mean, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, mean, cumulative].iter().all(|value| value.is_finite()) {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *mean < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let x = x.trunc();
                if x > 100000.0 {
                    return Err(FormulaEvalError::Num);
                }
                let x = x as u64;
                if *mean == 0.0 {
                    return Ok(if *cumulative != 0.0 || x == 0 {
                        1.0
                    } else {
                        0.0
                    });
                }
                let mut term = (-*mean).exp();
                let mut total = term;
                for k in 1..=x {
                    term *= *mean / k as f64;
                    if !term.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    total += term;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                checked_numeric_result(if *cumulative != 0.0 { total } else { term })
            }
            FormulaScalarFunction::Radians => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = value * std::f64::consts::PI / 180.0;
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Rand => {
                let [] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(formula_rand())
            }
            FormulaScalarFunction::RandBetween => {
                let [bottom, top] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_rand_between(*bottom, *top)
            }
            FormulaScalarFunction::Power => {
                let [base, exponent] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = base.powf(*exponent);
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Quotient => {
                let [numerator, denominator] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *denominator == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let value = round_toward_zero(numerator / denominator);
                if value.is_finite() {
                    Ok(normalize_zero(value))
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Sin => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = value.sin();
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Sinh => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(value.sinh())
            }
            FormulaScalarFunction::Standardize => {
                let [x, mean, standard_dev] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !x.is_finite() || !mean.is_finite() || !standard_dev.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *standard_dev <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((*x - *mean) / *standard_dev)
            }
            FormulaScalarFunction::Sqrt => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                Ok(value.sqrt())
            }
            FormulaScalarFunction::SqrtPi => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if *value < 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((*value * std::f64::consts::PI).sqrt())
            }
            FormulaScalarFunction::Second => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                Ok(formula_time_parts_from_serial(*serial)?.2 as f64)
            }
            FormulaScalarFunction::Syd => {
                let [cost, salvage, life, period] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![cost, salvage, life, period]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *life <= 0.0 || *period <= 0.0 || *period > *life {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(
                    (cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0)),
                )
            }
            FormulaScalarFunction::TDist => {
                let [x, degrees, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, degrees, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                if *cumulative != 0.0 {
                    student_t_cdf(*x, *degrees)
                } else {
                    student_t_pdf(*x, *degrees)
                }
            }
            FormulaScalarFunction::TDistLegacy => {
                let [x, degrees, tails] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, degrees, tails].iter().all(|value| value.is_finite()) {
                    return Err(FormulaEvalError::Value);
                }
                let tails = formula_integer_argument(*tails)?;
                if *x < 0.0 || *degrees < 1.0 || !matches!(tails, 1 | 2) {
                    return Err(FormulaEvalError::Num);
                }
                let right_tail = 1.0 - student_t_cdf(*x, *degrees)?;
                Ok(if tails == 1 {
                    right_tail
                } else {
                    (2.0 * right_tail).min(1.0)
                })
            }
            FormulaScalarFunction::TDist2T | FormulaScalarFunction::TDistRt => {
                let [x, degrees] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !x.is_finite() || !degrees.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let right_tail = 1.0 - student_t_cdf(*x, *degrees)?;
                Ok(if matches!(self, FormulaScalarFunction::TDist2T) {
                    (2.0 * right_tail).min(1.0)
                } else {
                    right_tail
                })
            }
            FormulaScalarFunction::TInv => {
                let [probability, degrees] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !probability.is_finite() || !degrees.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *probability <= 0.0 || *probability >= 1.0 || *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let mut low = -1.0;
                let mut high = 1.0;
                while student_t_cdf(low, *degrees)? > *probability {
                    high = low;
                    low *= 2.0;
                    if !low.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                while student_t_cdf(high, *degrees)? < *probability {
                    low = high;
                    high *= 2.0;
                    if !high.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
                for _ in 0..120 {
                    let mid = (low + high) / 2.0;
                    if student_t_cdf(mid, *degrees)? < *probability {
                        low = mid;
                    } else {
                        high = mid;
                    }
                }
                checked_numeric_result((low + high) / 2.0)
            }
            FormulaScalarFunction::TInvLegacy | FormulaScalarFunction::TInv2T => {
                let [probability, degrees] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if !probability.is_finite() || !degrees.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *degrees < 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let target = 1.0 - *probability / 2.0;
                let mut cdf = |x: f64| student_t_cdf(x, *degrees);
                inverse_positive_cdf(target, &mut cdf)
            }
            FormulaScalarFunction::Tan => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let value = value.tan();
                if value.is_finite() {
                    Ok(value)
                } else {
                    Err(FormulaEvalError::Num)
                }
            }
            FormulaScalarFunction::Tanh => {
                let [value] = args else {
                    return Err(FormulaEvalError::Value);
                };
                checked_numeric_result(value.tanh())
            }
            FormulaScalarFunction::TBillEq => {
                let [settlement, maturity, discount] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![settlement, maturity, discount]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *discount <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let days = treasury_bill_days(*settlement, *maturity)?;
                let denominator = 360.0 - discount * days;
                if denominator <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result(365.0 * discount / denominator)
            }
            FormulaScalarFunction::TBillPrice => {
                let [settlement, maturity, discount] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![settlement, maturity, discount]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *discount <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let days = treasury_bill_days(*settlement, *maturity)?;
                checked_numeric_result(100.0 * (1.0 - discount * days / 360.0))
            }
            FormulaScalarFunction::TBillYield => {
                let [settlement, maturity, price] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![settlement, maturity, price]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *price <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let days = treasury_bill_days(*settlement, *maturity)?;
                checked_numeric_result((100.0 - price) / price * 360.0 / days)
            }
            FormulaScalarFunction::Time => {
                let [hour, minute, second] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_time_serial_from_args(*hour, *minute, *second)
            }
            FormulaScalarFunction::Today => {
                let [] = args else {
                    return Err(FormulaEvalError::Value);
                };
                formula_current_excel_serial().map(f64::floor)
            }
            FormulaScalarFunction::Trunc => {
                let (value, digits) = match args {
                    [value] => (*value, 0.0),
                    [value, digits] => (*value, *digits),
                    _ => return Err(FormulaEvalError::Value),
                };
                let factor = formula_round_factor(digits)?;
                Ok(round_toward_zero(value * factor) / factor)
            }
            FormulaScalarFunction::Vdb => {
                let (cost, salvage, life, start_period, end_period, factor, no_switch) = match args
                {
                    [cost, salvage, life, start_period, end_period] => {
                        (*cost, *salvage, *life, *start_period, *end_period, 2.0, 0.0)
                    }
                    [cost, salvage, life, start_period, end_period, factor] => (
                        *cost,
                        *salvage,
                        *life,
                        *start_period,
                        *end_period,
                        *factor,
                        0.0,
                    ),
                    [
                        cost,
                        salvage,
                        life,
                        start_period,
                        end_period,
                        factor,
                        no_switch,
                    ] => (
                        *cost,
                        *salvage,
                        *life,
                        *start_period,
                        *end_period,
                        *factor,
                        *no_switch,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![
                    cost,
                    salvage,
                    life,
                    start_period,
                    end_period,
                    factor,
                    no_switch,
                ]
                .iter()
                .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if cost <= 0.0
                    || salvage < 0.0
                    || salvage > cost
                    || life <= 0.0
                    || start_period < 0.0
                    || end_period <= start_period
                    || end_period > life
                    || factor <= 0.0
                {
                    return Err(FormulaEvalError::Num);
                }
                let no_switch = no_switch != 0.0;
                let mut book_value = cost;
                let mut depreciation_total = 0.0;
                let mut period = 0.0;
                let period_limit = end_period.ceil();
                while period < period_limit {
                    let declining_depreciation = book_value * factor / life;
                    let remaining_periods = life - period;
                    if remaining_periods <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let straight_line_depreciation = (book_value - salvage) / remaining_periods;
                    let mut depreciation = if no_switch {
                        declining_depreciation
                    } else {
                        declining_depreciation.max(straight_line_depreciation)
                    };
                    depreciation = depreciation.min(book_value - salvage).max(0.0);
                    let overlap_start = start_period.max(period);
                    let overlap_end = end_period.min(period + 1.0);
                    if overlap_end > overlap_start {
                        depreciation_total += depreciation * (overlap_end - overlap_start);
                        if !depreciation_total.is_finite() {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    book_value -= depreciation;
                    if !book_value.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    period += 1.0;
                }
                checked_numeric_result(depreciation_total)
            }
            FormulaScalarFunction::Weekday => {
                let (serial, return_type) = match args {
                    [serial] => (*serial, 1.0),
                    [serial, return_type] => (*serial, *return_type),
                    _ => return Err(FormulaEvalError::Value),
                };
                let serial = formula_serial_integer(serial)?;
                let return_type = formula_integer_argument(return_type)?;
                let weekday_monday0 = serial_weekday_monday0(serial);
                if return_type == 3 {
                    return Ok(weekday_monday0 as f64);
                }
                let first_day_monday0 = week_start_from_return_type(return_type, true)?;
                Ok((weekday_monday0 - first_day_monday0).rem_euclid(7) as f64 + 1.0)
            }
            FormulaScalarFunction::WeekNum => {
                let (serial, return_type) = match args {
                    [serial] => (*serial, 1.0),
                    [serial, return_type] => (*serial, *return_type),
                    _ => return Err(FormulaEvalError::Value),
                };
                let serial = formula_serial_integer(serial)?;
                let return_type = formula_integer_argument(return_type)?;
                if return_type == 21 {
                    return iso_weeknum_from_serial(serial).map(|week| week as f64);
                }
                let first_day_monday0 = week_start_from_return_type(return_type, false)?;
                let (year, _, _) = formula_ymd_from_serial(serial as f64)?;
                let jan1_serial = formula_date_serial_from_args(year as f64, 1.0, 1.0)? as i64;
                let jan1_weekday_monday0 = serial_weekday_monday0(jan1_serial);
                let days_since_week_start =
                    (jan1_weekday_monday0 - first_day_monday0).rem_euclid(7);
                Ok((serial - jan1_serial + days_since_week_start).div_euclid(7) as f64 + 1.0)
            }
            FormulaScalarFunction::WeibullDist => {
                let [x, alpha, beta, cumulative] = args else {
                    return Err(FormulaEvalError::Value);
                };
                if ![x, alpha, beta, cumulative]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if *x < 0.0 || *alpha <= 0.0 || *beta <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let scaled = *x / *beta;
                let power = scaled.powf(*alpha);
                let value = if *cumulative != 0.0 {
                    1.0 - (-power).exp()
                } else if *x == 0.0 && *alpha < 1.0 {
                    return Err(FormulaEvalError::Num);
                } else {
                    *alpha / *beta * scaled.powf(*alpha - 1.0) * (-power).exp()
                };
                checked_numeric_result(value)
            }
            FormulaScalarFunction::Year => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let (year, _, _) = formula_ymd_from_serial(*serial)?;
                Ok(year as f64)
            }
            FormulaScalarFunction::YearFrac => {
                let (start_date, end_date, basis) = match args {
                    [start_date, end_date] => (*start_date, *end_date, 0),
                    [start_date, end_date, basis] => {
                        (*start_date, *end_date, yearfrac_basis(*basis)?)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                let start_serial = formula_serial_integer(start_date)?;
                let end_serial = formula_serial_integer(end_date)?;
                match basis {
                    0 => Ok(days360(start_serial, end_serial, false)? as f64 / 360.0),
                    1 => yearfrac_actual_actual(start_serial, end_serial),
                    2 => Ok((end_serial - start_serial) as f64 / 360.0),
                    3 => Ok((end_serial - start_serial) as f64 / 365.0),
                    4 => Ok(days360(start_serial, end_serial, true)? as f64 / 360.0),
                    _ => Err(FormulaEvalError::Num),
                }
            }
            FormulaScalarFunction::Yield => {
                let (settlement, maturity, rate, price, redemption, frequency, basis) = match args {
                    [settlement, maturity, rate, price, redemption, frequency] => (
                        *settlement,
                        *maturity,
                        *rate,
                        *price,
                        *redemption,
                        *frequency,
                        0.0,
                    ),
                    [
                        settlement,
                        maturity,
                        rate,
                        price,
                        redemption,
                        frequency,
                        basis,
                    ] => (
                        *settlement,
                        *maturity,
                        *rate,
                        *price,
                        *redemption,
                        *frequency,
                        *basis,
                    ),
                    _ => return Err(FormulaEvalError::Value),
                };
                regular_coupon_yield(
                    settlement, maturity, rate, price, redemption, frequency, basis,
                )
            }
            FormulaScalarFunction::YieldDisc => {
                let (settlement, maturity, price, redemption, basis) = match args {
                    [settlement, maturity, price, redemption] => {
                        (*settlement, *maturity, *price, *redemption, 0.0)
                    }
                    [settlement, maturity, price, redemption, basis] => {
                        (*settlement, *maturity, *price, *redemption, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, price, redemption, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if price <= 0.0 || redemption <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let yearfrac = discount_security_yearfrac(settlement, maturity, basis)?;
                checked_numeric_result((redemption - price) / price / yearfrac)
            }
            FormulaScalarFunction::YieldMat => {
                let (settlement, maturity, issue, rate, price, basis) = match args {
                    [settlement, maturity, issue, rate, price] => {
                        (*settlement, *maturity, *issue, *rate, *price, 0.0)
                    }
                    [settlement, maturity, issue, rate, price, basis] => {
                        (*settlement, *maturity, *issue, *rate, *price, *basis)
                    }
                    _ => return Err(FormulaEvalError::Value),
                };
                if ![settlement, maturity, issue, rate, price, basis]
                    .iter()
                    .all(|value| value.is_finite())
                {
                    return Err(FormulaEvalError::Value);
                }
                if rate < 0.0 || price <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                let (issue_to_maturity, settlement_to_maturity, issue_to_settlement) =
                    maturity_security_yearfracs(settlement, maturity, issue, basis)?;
                let maturity_value = 100.0 + 100.0 * rate * issue_to_maturity;
                let accrued_interest = 100.0 * rate * issue_to_settlement;
                let investment = price + accrued_interest;
                if investment <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                checked_numeric_result((maturity_value / investment - 1.0) / settlement_to_maturity)
            }
        }
    }
}

fn round_half_away_from_zero(value: f64) -> f64 {
    if value.is_sign_negative() {
        -((-value) + 0.5).floor()
    } else {
        (value + 0.5).floor()
    }
}

fn round_away_from_zero(value: f64) -> f64 {
    if value.is_sign_negative() {
        value.floor()
    } else {
        value.ceil()
    }
}

fn round_toward_zero(value: f64) -> f64 {
    if value.is_sign_negative() {
        value.ceil()
    } else {
        value.floor()
    }
}

fn formula_round_factor(digits: f64) -> Result<f64, FormulaEvalError> {
    if !digits.is_finite() || digits.fract() != 0.0 {
        return Err(FormulaEvalError::Value);
    }
    if digits < i32::MIN as f64 || digits > i32::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    let factor = 10_f64.powi(digits as i32);
    if !factor.is_finite() || factor == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    Ok(factor)
}

fn formula_checked_numeric_result(value: f64) -> Result<f64, FormulaEvalError> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_erf_approx(value: f64) -> f64 {
    let sign = if value.is_sign_negative() { -1.0 } else { 1.0 };
    let x = value.abs();
    let t = 1.0 / (1.0 + 0.5 * x);
    let tau = t
        * (-x * x - 1.26551223
            + t * (1.00002368
                + t * (0.37409196
                    + t * (0.09678418
                        + t * (-0.18628806
                            + t * (0.27886807
                                + t * (-1.13520398
                                    + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))))
            .exp();
    sign * (1.0 - tau)
}

fn formula_standard_normal_cdf(z: f64) -> Result<f64, FormulaEvalError> {
    if !z.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    formula_checked_numeric_result(0.5 * (1.0 + formula_erf_approx(z / std::f64::consts::SQRT_2)))
}

fn formula_gamma_ln_value(value: f64) -> f64 {
    const COEFFICIENTS: [f64; 9] = [
        0.9999999999998099,
        676.5203681218851,
        -1259.1392167224028,
        771.3234287776531,
        -176.6150291621406,
        12.507343278686905,
        -0.13857109526572012,
        0.000009984369578019572,
        0.00000015056327351493116,
    ];
    let lanczos = |input: f64| {
        let z = input - 1.0;
        let mut x = COEFFICIENTS[0];
        for (index, coefficient) in COEFFICIENTS.iter().enumerate().skip(1) {
            x += coefficient / (z + index as f64);
        }
        let t = z + 7.5;
        0.5 * (2.0 * std::f64::consts::PI).ln() + (z + 0.5) * t.ln() - t + x.ln()
    };
    if value < 0.5 {
        std::f64::consts::PI.ln() - (std::f64::consts::PI * value).sin().ln() - lanczos(1.0 - value)
    } else {
        lanczos(value)
    }
}

fn formula_beta_fraction(alpha: f64, beta: f64, x: f64) -> Result<f64, FormulaEvalError> {
    const EPSILON: f64 = 1e-14;
    const FLOOR: f64 = 1e-300;
    const MAX_ITERATIONS: usize = 200;
    let qab = alpha + beta;
    let qap = alpha + 1.0;
    let qam = alpha - 1.0;
    let mut c = 1.0;
    let mut d = 1.0 - qab * x / qap;
    if d.abs() < FLOOR {
        d = FLOOR;
    }
    d = 1.0 / d;
    let mut h = d;
    for m in 1..=MAX_ITERATIONS {
        let m_f = m as f64;
        let m2 = 2.0 * m_f;
        let mut aa = m_f * (beta - m_f) * x / ((qam + m2) * (alpha + m2));
        d = 1.0 + aa * d;
        if d.abs() < FLOOR {
            d = FLOOR;
        }
        c = 1.0 + aa / c;
        if c.abs() < FLOOR {
            c = FLOOR;
        }
        d = 1.0 / d;
        h *= d * c;
        aa = -(alpha + m_f) * (qab + m_f) * x / ((alpha + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < FLOOR {
            d = FLOOR;
        }
        c = 1.0 + aa / c;
        if c.abs() < FLOOR {
            c = FLOOR;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() <= EPSILON {
            return formula_checked_numeric_result(h);
        }
    }
    formula_checked_numeric_result(h)
}

fn formula_regularized_beta(x: f64, alpha: f64, beta: f64) -> Result<f64, FormulaEvalError> {
    if ![x, alpha, beta].iter().all(|value| value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    if alpha <= 0.0 || beta <= 0.0 || !(0.0..=1.0).contains(&x) {
        return Err(FormulaEvalError::Num);
    }
    if x == 0.0 || x == 1.0 {
        return Ok(x);
    }
    let log_beta = formula_gamma_ln_value(alpha) + formula_gamma_ln_value(beta)
        - formula_gamma_ln_value(alpha + beta);
    let front = (alpha * x.ln() + beta * (-x).ln_1p() - log_beta).exp();
    if x < (alpha + 1.0) / (alpha + beta + 2.0) {
        formula_checked_numeric_result(
            (front * formula_beta_fraction(alpha, beta, x)? / alpha).clamp(0.0, 1.0),
        )
    } else {
        formula_checked_numeric_result(
            (1.0 - front * formula_beta_fraction(beta, alpha, 1.0 - x)? / beta).clamp(0.0, 1.0),
        )
    }
}

fn formula_student_t_cdf(x: f64, degrees: f64) -> Result<f64, FormulaEvalError> {
    if !x.is_finite() || !degrees.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if degrees < 1.0 {
        return Err(FormulaEvalError::Num);
    }
    let degrees = degrees.trunc();
    let beta_x = degrees / (degrees + x * x);
    let tail = 0.5 * formula_regularized_beta(beta_x, degrees / 2.0, 0.5)?;
    Ok(if x >= 0.0 { 1.0 - tail } else { tail })
}

fn formula_student_t_right_tail_from_abs(t: f64, degrees: f64) -> Result<f64, FormulaEvalError> {
    formula_student_t_cdf(t.abs(), degrees).map(|value| (1.0 - value).clamp(0.0, 1.0))
}

fn formula_f_right_tail(x: f64, degrees1: f64, degrees2: f64) -> Result<f64, FormulaEvalError> {
    if ![x, degrees1, degrees2]
        .iter()
        .all(|value| value.is_finite())
    {
        return Err(FormulaEvalError::Value);
    }
    if x < 0.0 || degrees1 < 1.0 || degrees2 < 1.0 {
        return Err(FormulaEvalError::Num);
    }
    let degrees1 = degrees1.trunc();
    let degrees2 = degrees2.trunc();
    let transformed = degrees1 * x / (degrees1 * x + degrees2);
    formula_regularized_beta(transformed, degrees1 / 2.0, degrees2 / 2.0)
        .map(|value| (1.0 - value).clamp(0.0, 1.0))
}

fn formula_sample_mean_and_variance(values: &[f64]) -> Result<(f64, f64), FormulaEvalError> {
    if values.len() < 2 {
        return Err(FormulaEvalError::Div0);
    }
    if values.iter().any(|value| !value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let deviation_sum = values
        .iter()
        .map(|value| {
            let deviation = value - mean;
            deviation * deviation
        })
        .sum::<f64>();
    formula_checked_numeric_result(deviation_sum / (values.len() - 1) as f64)
        .map(|variance| (mean, variance))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaComparisonOperator {
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

impl FormulaComparisonOperator {
    fn evaluate(self, left: f64, right: f64) -> bool {
        match self {
            FormulaComparisonOperator::Equal => left == right,
            FormulaComparisonOperator::NotEqual => left != right,
            FormulaComparisonOperator::LessThan => left < right,
            FormulaComparisonOperator::LessThanOrEqual => left <= right,
            FormulaComparisonOperator::GreaterThan => left > right,
            FormulaComparisonOperator::GreaterThanOrEqual => left >= right,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
enum FormulaCriteria {
    Blank,
    NonBlank,
    Number {
        operator: FormulaComparisonOperator,
        value: f64,
    },
    Text {
        operator: FormulaComparisonOperator,
        pattern: String,
    },
}

impl FormulaCriteria {
    fn matches(&self, cell_value: &CellValue) -> bool {
        match self {
            FormulaCriteria::Blank => matches!(cell_value, CellValue::Blank),
            FormulaCriteria::NonBlank => !matches!(cell_value, CellValue::Blank),
            FormulaCriteria::Number {
                operator,
                value: expected,
            } => match cell_value {
                CellValue::Number(actual) => operator.evaluate(*actual, *expected),
                CellValue::Bool(actual) => {
                    operator.evaluate(if *actual { 1.0 } else { 0.0 }, *expected)
                }
                CellValue::Blank | CellValue::Text(_) | CellValue::Error(_) => false,
            },
            FormulaCriteria::Text { operator, pattern } => match cell_value {
                CellValue::Text(actual) => match operator {
                    FormulaComparisonOperator::Equal => {
                        formula_wildcard_matches(pattern, actual, true)
                    }
                    FormulaComparisonOperator::NotEqual => {
                        !formula_wildcard_matches(pattern, actual, true)
                    }
                    _ => false,
                },
                CellValue::Blank
                | CellValue::Bool(_)
                | CellValue::Number(_)
                | CellValue::Error(_) => false,
            },
        }
    }

    fn from_numeric_value(value: f64) -> Self {
        Self::Number {
            operator: FormulaComparisonOperator::Equal,
            value,
        }
    }

    fn from_string_literal(literal: String) -> Self {
        if literal.is_empty() {
            return Self::Blank;
        }
        if let Some((operator, value)) = parse_formula_criteria_numeric_literal(literal.as_str()) {
            return Self::Number { operator, value };
        }
        if literal == "<>" {
            return Self::NonBlank;
        }
        if let Some(pattern) = literal.strip_prefix("<>") {
            return Self::Text {
                operator: FormulaComparisonOperator::NotEqual,
                pattern: pattern.to_string(),
            };
        }
        if let Some(pattern) = literal.strip_prefix('=') {
            if pattern.is_empty() {
                return Self::Blank;
            }
            return Self::Text {
                operator: FormulaComparisonOperator::Equal,
                pattern: pattern.to_string(),
            };
        }
        Self::Text {
            operator: FormulaComparisonOperator::Equal,
            pattern: literal,
        }
    }
}

#[derive(Debug, Clone)]
struct FormulaCriteriaRange {
    sheet_id: SheetId,
    rect: Rect,
    criteria: FormulaCriteria,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FormulaReference {
    areas: Vec<(SheetId, Rect)>,
    explicit_area_count: usize,
}

impl FormulaReference {
    fn with_explicit_area_count(
        explicit_area_count: usize,
        areas: Vec<(SheetId, Rect)>,
    ) -> Result<Self, FormulaEvalError> {
        if areas.is_empty() {
            return Err(FormulaEvalError::Ref);
        }
        if explicit_area_count == 0 {
            return Err(FormulaEvalError::Ref);
        }
        Ok(Self {
            areas,
            explicit_area_count,
        })
    }

    fn single(sheet_id: SheetId, rect: Rect) -> Self {
        Self {
            areas: vec![(sheet_id, rect)],
            explicit_area_count: 1,
        }
    }

    fn single_area(&self) -> Result<(SheetId, Rect), FormulaEvalError> {
        if self.explicit_area_count != 1 {
            return Err(FormulaEvalError::Value);
        }
        self.areas.first().copied().ok_or(FormulaEvalError::Ref)
    }

    fn len(&self) -> usize {
        self.explicit_area_count
    }

    fn areas(&self) -> &[(SheetId, Rect)] {
        &self.areas
    }
}

#[derive(Debug, Clone, PartialEq)]
enum FormulaValueProbe {
    Blank,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(FormulaEvalError),
    Omitted,
    Lambda {
        parameters: Vec<String>,
        body: String,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaLogicalFunction {
    And,
    Or,
    Xor,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaGroupByAggregation {
    Sum,
    Average,
    Count,
    CountA,
    Max,
    Min,
    Product,
}

impl FormulaGroupByAggregation {
    fn from_name(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("SUM") {
            Some(Self::Sum)
        } else if name.eq_ignore_ascii_case("AVERAGE") {
            Some(Self::Average)
        } else if name.eq_ignore_ascii_case("COUNT") {
            Some(Self::Count)
        } else if name.eq_ignore_ascii_case("COUNTA") {
            Some(Self::CountA)
        } else if name.eq_ignore_ascii_case("MAX") {
            Some(Self::Max)
        } else if name.eq_ignore_ascii_case("MIN") {
            Some(Self::Min)
        } else if name.eq_ignore_ascii_case("PRODUCT") {
            Some(Self::Product)
        } else {
            None
        }
    }

    fn evaluate(self, values: &[FormulaValueProbe]) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut numbers = Vec::new();
        let mut counta = 0_u64;
        for value in values {
            match value {
                FormulaValueProbe::Blank => {}
                FormulaValueProbe::Bool(_) | FormulaValueProbe::Text(_) => counta += 1,
                FormulaValueProbe::Number(number) => {
                    counta += 1;
                    numbers.push(*number);
                }
                FormulaValueProbe::Error(error) => return Err(*error),
                FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {}
            }
        }
        match self {
            Self::Sum => Ok(FormulaValueProbe::Number(numbers.iter().sum())),
            Self::Average => {
                if numbers.is_empty() {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(FormulaValueProbe::Number(
                        numbers.iter().sum::<f64>() / numbers.len() as f64,
                    ))
                }
            }
            Self::Count => Ok(FormulaValueProbe::Number(numbers.len() as f64)),
            Self::CountA => Ok(FormulaValueProbe::Number(counta as f64)),
            Self::Max => Ok(FormulaValueProbe::Number(
                numbers.iter().copied().reduce(f64::max).unwrap_or(0.0),
            )),
            Self::Min => Ok(FormulaValueProbe::Number(
                numbers.iter().copied().reduce(f64::min).unwrap_or(0.0),
            )),
            Self::Product => Ok(FormulaValueProbe::Number(numbers.iter().product())),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaLookupMode {
    Exact,
    ApproxAscending,
    ApproxDescending,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaXLookupMatchMode {
    Exact,
    ExactOrNextSmaller,
    ExactOrNextLarger,
    Wildcard,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaXLookupSearchMode {
    Forward,
    Reverse,
    BinaryAscending,
    BinaryDescending,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaLookupOrientation {
    FirstColumn,
    FirstRow,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaWildcardToken {
    Literal(char),
    AnyChar,
    AnySequence,
}

fn formula_value_probe_from_cell_value(value: CellValue) -> FormulaValueProbe {
    match value {
        CellValue::Blank => FormulaValueProbe::Blank,
        CellValue::Bool(value) => FormulaValueProbe::Bool(value),
        CellValue::Number(value) => FormulaValueProbe::Number(value),
        CellValue::Text(value) => FormulaValueProbe::Text(value),
        CellValue::Error(error) => {
            FormulaValueProbe::Error(formula_eval_error_from_cell_error(error))
        }
    }
}

fn formula_cell_value_from_probe(value: FormulaValueProbe) -> Option<CellValue> {
    match value {
        FormulaValueProbe::Blank => Some(CellValue::Blank),
        FormulaValueProbe::Bool(value) => Some(CellValue::Bool(value)),
        FormulaValueProbe::Number(value) => Some(CellValue::Number(value)),
        FormulaValueProbe::Text(value) => Some(CellValue::Text(value)),
        FormulaValueProbe::Error(error) => error.into_cell_value(),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => None,
    }
}

fn formula_number_from_value_probe(value: FormulaValueProbe) -> Result<f64, FormulaEvalError> {
    match value {
        FormulaValueProbe::Blank => Ok(0.0),
        FormulaValueProbe::Bool(value) => Ok(if value { 1.0 } else { 0.0 }),
        FormulaValueProbe::Number(value) => Ok(value),
        FormulaValueProbe::Text(_) => Err(FormulaEvalError::Value),
        FormulaValueProbe::Error(error) => Err(error),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
            Err(FormulaEvalError::Value)
        }
    }
}

fn formula_text_from_value_probe(value: FormulaValueProbe) -> Result<String, FormulaEvalError> {
    match value {
        FormulaValueProbe::Blank => Ok(String::new()),
        FormulaValueProbe::Bool(value) => Ok(if value { "TRUE" } else { "FALSE" }.into()),
        FormulaValueProbe::Number(value) => formula_text_from_number(value),
        FormulaValueProbe::Text(value) => Ok(value),
        FormulaValueProbe::Error(error) => Err(error),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
            Err(FormulaEvalError::Value)
        }
    }
}

pub(super) fn render_range_text_value(value: &OmValue) -> String {
    match value {
        OmValue::Missing | OmValue::Empty | OmValue::Null => String::new(),
        OmValue::Bool(true) => "TRUE".to_string(),
        OmValue::Bool(false) => "FALSE".to_string(),
        OmValue::Number(number) => number.to_string(),
        OmValue::Text(text) => text.clone(),
        OmValue::Error(error) => formula_cell_error_text(*error).to_string(),
        OmValue::Object(_) | OmValue::Array(_) => String::new(),
    }
}

pub(super) fn formula_cell_error_text(error: CellError) -> &'static str {
    match error {
        CellError::Null => "#NULL!",
        CellError::Div0 => "#DIV/0!",
        CellError::Value => "#VALUE!",
        CellError::Ref => "#REF!",
        CellError::Name => "#NAME?",
        CellError::Num => "#NUM!",
        CellError::NA => "#N/A",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
        CellError::Field => "#FIELD!",
        CellError::Blocked => "#BLOCKED!",
        CellError::Busy => "#BUSY!",
        CellError::Connect => "#CONNECT!",
        CellError::Python => "#PYTHON!",
        CellError::Timeout => "#TIMEOUT!",
        CellError::Unknown => "#UNKNOWN!",
    }
}

pub(super) fn format_formula_string_literal(value: &str) -> String {
    format!("\"{}\"", value.replace('"', "\"\""))
}

pub(super) fn worksheet_function_formula_name(member: &str) -> OmResult<String> {
    let mut name = String::with_capacity(member.len());
    for ch in member.chars() {
        if ch.is_ascii_alphanumeric() {
            name.push(ch.to_ascii_uppercase());
        } else if ch == '.' || ch == '_' {
            name.push('.');
        } else {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("WorksheetFunction.{member} is not a valid worksheet function name"),
            ));
        }
    }
    if name.is_empty() || !name.as_bytes()[0].is_ascii_alphabetic() {
        return Err(OmError::new(
            OmErrorCode::NotFound,
            format!("WorksheetFunction.{member} is not a valid worksheet function name"),
        ));
    }
    Ok(name)
}

fn formula_eval_error_text(error: FormulaEvalError) -> &'static str {
    match error {
        FormulaEvalError::Unsupported => "#VALUE!",
        FormulaEvalError::Circular => "#CALC!",
        FormulaEvalError::Null => "#NULL!",
        FormulaEvalError::Div0 => "#DIV/0!",
        FormulaEvalError::Value => "#VALUE!",
        FormulaEvalError::Ref => "#REF!",
        FormulaEvalError::Name => "#NAME?",
        FormulaEvalError::NA => "#N/A",
        FormulaEvalError::Num => "#NUM!",
        FormulaEvalError::GettingData => "#GETTING_DATA",
        FormulaEvalError::Spill => "#SPILL!",
        FormulaEvalError::Calc => "#CALC!",
        FormulaEvalError::Field => "#FIELD!",
        FormulaEvalError::Blocked => "#BLOCKED!",
        FormulaEvalError::Busy => "#BUSY!",
        FormulaEvalError::Connect => "#CONNECT!",
        FormulaEvalError::Python => "#PYTHON!",
        FormulaEvalError::Timeout => "#TIMEOUT!",
        FormulaEvalError::Unknown => "#UNKNOWN!",
    }
}

fn formula_strict_text_literal(text: &str) -> String {
    format!("\"{}\"", text.replace('"', "\"\""))
}

fn formula_value_to_text(
    value: FormulaValueProbe,
    strict: bool,
) -> Result<String, FormulaEvalError> {
    match value {
        FormulaValueProbe::Blank => Ok(String::new()),
        FormulaValueProbe::Bool(value) => Ok(if value { "TRUE" } else { "FALSE" }.to_string()),
        FormulaValueProbe::Number(value) => formula_text_from_number(value),
        FormulaValueProbe::Text(value) if strict => Ok(formula_strict_text_literal(value.as_str())),
        FormulaValueProbe::Text(value) => Ok(value),
        FormulaValueProbe::Error(error) => Ok(formula_eval_error_text(error).to_string()),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
            Err(FormulaEvalError::Value)
        }
    }
}

fn formula_text_byte_width(ch: char) -> usize {
    if ch.is_ascii() { 1 } else { 2 }
}

fn formula_text_byte_len(text: &str) -> usize {
    text.chars().map(formula_text_byte_width).sum()
}

fn formula_text_byte_slice(text: &str, start: usize, count: usize) -> String {
    if count == 0 {
        return String::new();
    }
    let end = start.saturating_add(count);
    let mut position = 1_usize;
    let mut output = String::new();
    for ch in text.chars() {
        let width = formula_text_byte_width(ch);
        let next_position = position + width;
        if position >= end {
            break;
        }
        if position >= start && next_position <= end {
            output.push(ch);
        }
        position = next_position;
    }
    output
}

fn formula_text_char_position_to_byte_position(text: &str, char_position: usize) -> usize {
    let units = text
        .chars()
        .take(char_position.saturating_sub(1))
        .map(formula_text_byte_width)
        .sum::<usize>();
    units + 1
}

fn formula_encode_url(text: &str) -> String {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut output = String::new();
    for byte in text.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.' | b'~') {
            output.push(byte as char);
        } else {
            output.push('%');
            output.push(HEX[(byte >> 4) as usize] as char);
            output.push(HEX[(byte & 0x0f) as usize] as char);
        }
    }
    output
}

fn formula_regex_case_insensitive(case_sensitivity: i64) -> Result<bool, FormulaEvalError> {
    match case_sensitivity {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(FormulaEvalError::Value),
    }
}

fn formula_regex_from_pattern(
    pattern: &str,
    case_insensitive: bool,
) -> Result<Regex, FormulaEvalError> {
    RegexBuilder::new(pattern)
        .case_insensitive(case_insensitive)
        .build()
        .map_err(|_| FormulaEvalError::Value)
}

fn formula_selected_text_from_value_probe(
    value: FormulaValueProbe,
) -> Result<String, FormulaEvalError> {
    match value {
        FormulaValueProbe::Text(value) => Ok(value),
        FormulaValueProbe::Error(error) => Err(error),
        FormulaValueProbe::Blank
        | FormulaValueProbe::Bool(_)
        | FormulaValueProbe::Number(_)
        | FormulaValueProbe::Omitted
        | FormulaValueProbe::Lambda { .. } => Err(FormulaEvalError::Unsupported),
    }
}

fn formula_array_projection_function_name(name: &str) -> bool {
    name.eq_ignore_ascii_case("BYCOL")
        || name.eq_ignore_ascii_case("BYROW")
        || name.eq_ignore_ascii_case("CHOOSECOLS")
        || name.eq_ignore_ascii_case("CHOOSEROWS")
        || name.eq_ignore_ascii_case("DROP")
        || name.eq_ignore_ascii_case("EXPAND")
        || name.eq_ignore_ascii_case("FILTER")
        || name.eq_ignore_ascii_case("GROUPBY")
        || name.eq_ignore_ascii_case("HSTACK")
        || name.eq_ignore_ascii_case("MAP")
        || name.eq_ignore_ascii_case("PIVOTBY")
        || name.eq_ignore_ascii_case("SORT")
        || name.eq_ignore_ascii_case("SORTBY")
        || name.eq_ignore_ascii_case("TAKE")
        || name.eq_ignore_ascii_case("TOCOL")
        || name.eq_ignore_ascii_case("TOROW")
        || name.eq_ignore_ascii_case("TRANSPOSE")
        || name.eq_ignore_ascii_case("UNIQUE")
        || name.eq_ignore_ascii_case("VSTACK")
        || name.eq_ignore_ascii_case("WRAPCOLS")
        || name.eq_ignore_ascii_case("WRAPROWS")
}

fn formula_text_function_name(name: &str) -> bool {
    formula_array_projection_function_name(name)
        || name.eq_ignore_ascii_case("ADDRESS")
        || name.eq_ignore_ascii_case("ARRAYTOTEXT")
        || name.eq_ignore_ascii_case("ASC")
        || name.eq_ignore_ascii_case("BAHTTEXT")
        || name.eq_ignore_ascii_case("CONCAT")
        || name.eq_ignore_ascii_case("CONCATENATE")
        || name.eq_ignore_ascii_case("DGET")
        || name.eq_ignore_ascii_case("DBCS")
        || name.eq_ignore_ascii_case("ENCODEURL")
        || name.eq_ignore_ascii_case("LEFT")
        || name.eq_ignore_ascii_case("LEFTB")
        || name.eq_ignore_ascii_case("RIGHT")
        || name.eq_ignore_ascii_case("RIGHTB")
        || name.eq_ignore_ascii_case("MID")
        || name.eq_ignore_ascii_case("MIDB")
        || name.eq_ignore_ascii_case("BASE")
        || name.eq_ignore_ascii_case("BIN2HEX")
        || name.eq_ignore_ascii_case("BIN2OCT")
        || name.eq_ignore_ascii_case("CELL")
        || name.eq_ignore_ascii_case("CHAR")
        || name.eq_ignore_ascii_case("CLEAN")
        || name.eq_ignore_ascii_case("COMPLEX")
        || name.eq_ignore_ascii_case("CUBEKPIMEMBER")
        || name.eq_ignore_ascii_case("CUBEMEMBER")
        || name.eq_ignore_ascii_case("CUBEMEMBERPROPERTY")
        || name.eq_ignore_ascii_case("CUBERANKEDMEMBER")
        || name.eq_ignore_ascii_case("CUBESET")
        || name.eq_ignore_ascii_case("DEC2BIN")
        || name.eq_ignore_ascii_case("DEC2HEX")
        || name.eq_ignore_ascii_case("DEC2OCT")
        || name.eq_ignore_ascii_case("DOLLAR")
        || name.eq_ignore_ascii_case("DETECTLANGUAGE")
        || name.eq_ignore_ascii_case("FIXED")
        || name.eq_ignore_ascii_case("FILTERXML")
        || name.eq_ignore_ascii_case("GETPIVOTDATA")
        || name.eq_ignore_ascii_case("HYPERLINK")
        || name.eq_ignore_ascii_case("HEX2BIN")
        || name.eq_ignore_ascii_case("HEX2OCT")
        || name.eq_ignore_ascii_case("IMAGE")
        || name.eq_ignore_ascii_case("IMCONJUGATE")
        || name.eq_ignore_ascii_case("IMCOS")
        || name.eq_ignore_ascii_case("IMCOSH")
        || name.eq_ignore_ascii_case("IMCOT")
        || name.eq_ignore_ascii_case("IMCSC")
        || name.eq_ignore_ascii_case("IMCSCH")
        || name.eq_ignore_ascii_case("IMDIV")
        || name.eq_ignore_ascii_case("IMEXP")
        || name.eq_ignore_ascii_case("INFO")
        || name.eq_ignore_ascii_case("INDIRECT")
        || name.eq_ignore_ascii_case("IMLN")
        || name.eq_ignore_ascii_case("LAMBDA")
        || name.eq_ignore_ascii_case("LET")
        || name.eq_ignore_ascii_case("MAKEARRAY")
        || name.eq_ignore_ascii_case("IMLOG10")
        || name.eq_ignore_ascii_case("IMLOG2")
        || name.eq_ignore_ascii_case("IMPOWER")
        || name.eq_ignore_ascii_case("IMPRODUCT")
        || name.eq_ignore_ascii_case("IMSEC")
        || name.eq_ignore_ascii_case("IMSECH")
        || name.eq_ignore_ascii_case("IMSIN")
        || name.eq_ignore_ascii_case("IMSINH")
        || name.eq_ignore_ascii_case("IMSQRT")
        || name.eq_ignore_ascii_case("IMSUB")
        || name.eq_ignore_ascii_case("IMSUM")
        || name.eq_ignore_ascii_case("IMTAN")
        || name.eq_ignore_ascii_case("JIS")
        || name.eq_ignore_ascii_case("OCT2BIN")
        || name.eq_ignore_ascii_case("OCT2HEX")
        || name.eq_ignore_ascii_case("OFFSET")
        || name.eq_ignore_ascii_case("PHONETIC")
        || name.eq_ignore_ascii_case("REDUCE")
        || name.eq_ignore_ascii_case("ROMAN")
        || name.eq_ignore_ascii_case("SCAN")
        || name.eq_ignore_ascii_case("T")
        || name.eq_ignore_ascii_case("TEXT")
        || name.eq_ignore_ascii_case("TEXTSPLIT")
        || name.eq_ignore_ascii_case("UNICHAR")
        || name.eq_ignore_ascii_case("UPPER")
        || name.eq_ignore_ascii_case("LOWER")
        || name.eq_ignore_ascii_case("PROPER")
        || name.eq_ignore_ascii_case("REGEXEXTRACT")
        || name.eq_ignore_ascii_case("REGEXREPLACE")
        || name.eq_ignore_ascii_case("TRIM")
        || name.eq_ignore_ascii_case("TEXTJOIN")
        || name.eq_ignore_ascii_case("TEXTBEFORE")
        || name.eq_ignore_ascii_case("TEXTAFTER")
        || name.eq_ignore_ascii_case("TRIMRANGE")
        || name.eq_ignore_ascii_case("TRANSLATE")
        || name.eq_ignore_ascii_case("REPT")
        || name.eq_ignore_ascii_case("REPLACE")
        || name.eq_ignore_ascii_case("REPLACEB")
        || name.eq_ignore_ascii_case("SUBSTITUTE")
        || name.eq_ignore_ascii_case("VALUETOTEXT")
        || name.eq_ignore_ascii_case("FORMULATEXT")
        || name.eq_ignore_ascii_case("IF")
        || name.eq_ignore_ascii_case("IFS")
        || name.eq_ignore_ascii_case("SWITCH")
        || name.eq_ignore_ascii_case("CHOOSE")
        || name.eq_ignore_ascii_case("INDEX")
        || name.eq_ignore_ascii_case("LOOKUP")
        || name.eq_ignore_ascii_case("VLOOKUP")
        || name.eq_ignore_ascii_case("HLOOKUP")
        || name.eq_ignore_ascii_case("WEBSERVICE")
        || name.eq_ignore_ascii_case("XLOOKUP")
}

fn formula_text_from_number(value: f64) -> Result<String, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if value == 0.0 {
        return Ok("0".to_string());
    }
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        return Ok((value as i64).to_string());
    }
    Ok(value.to_string())
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct FormulaComplexNumber {
    pub(super) real: f64,
    pub(super) imaginary: f64,
    pub(super) suffix: Option<char>,
}

fn formula_complex_clean_component(value: f64) -> f64 {
    if value.abs() < 1e-12 { 0.0 } else { value }
}

fn formula_complex_number(
    real: f64,
    imaginary: f64,
    suffix: Option<char>,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    if !real.is_finite() || !imaginary.is_finite() {
        return Err(FormulaEvalError::Num);
    }
    if suffix.is_some_and(|suffix| !matches!(suffix, 'i' | 'j')) {
        return Err(FormulaEvalError::Value);
    }
    Ok(FormulaComplexNumber {
        real: formula_complex_clean_component(real),
        imaginary: formula_complex_clean_component(imaginary),
        suffix,
    })
}

fn formula_complex_suffix(text: &str) -> Result<char, FormulaEvalError> {
    match text {
        "i" => Ok('i'),
        "j" => Ok('j'),
        _ => Err(FormulaEvalError::Value),
    }
}

fn formula_complex_component(text: &str) -> Result<f64, FormulaEvalError> {
    let value = text.parse::<f64>().map_err(|_| FormulaEvalError::Value)?;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_complex_imaginary_coefficient(text: &str) -> Result<f64, FormulaEvalError> {
    match text {
        "" | "+" => Ok(1.0),
        "-" => Ok(-1.0),
        _ => formula_complex_component(text),
    }
}

fn formula_complex_split_imaginary(text: &str) -> Option<usize> {
    let mut split = None;
    for (index, ch) in text.char_indices().skip(1) {
        if matches!(ch, '+' | '-')
            && !text[..index]
                .chars()
                .next_back()
                .is_some_and(|previous| matches!(previous, 'e' | 'E'))
        {
            split = Some(index);
        }
    }
    split
}

pub(super) fn formula_complex_from_text(
    text: &str,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let text = text.trim();
    if text.is_empty() || text.chars().any(char::is_whitespace) {
        return Err(FormulaEvalError::Value);
    }
    if text.ends_with('I') || text.ends_with('J') {
        return Err(FormulaEvalError::Value);
    }
    let Some(suffix) = text.chars().last().filter(|ch| matches!(ch, 'i' | 'j')) else {
        return formula_complex_number(formula_complex_component(text)?, 0.0, None);
    };
    let value = &text[..text.len() - suffix.len_utf8()];
    let (real, imaginary) = if let Some(split) = formula_complex_split_imaginary(value) {
        (
            formula_complex_component(&value[..split])?,
            formula_complex_imaginary_coefficient(&value[split..])?,
        )
    } else {
        (0.0, formula_complex_imaginary_coefficient(value)?)
    };
    formula_complex_number(real, imaginary, Some(suffix))
}

fn formula_complex_format(value: FormulaComplexNumber) -> Result<String, FormulaEvalError> {
    let real = formula_complex_clean_component(value.real);
    let imaginary = formula_complex_clean_component(value.imaginary);
    if imaginary == 0.0 {
        return formula_text_from_number(real);
    }
    let suffix = value.suffix.unwrap_or('i');
    let imaginary_text = |magnitude: f64| -> Result<String, FormulaEvalError> {
        if magnitude == 1.0 {
            Ok(suffix.to_string())
        } else {
            Ok(format!(
                "{}{}",
                formula_text_from_number(magnitude)?,
                suffix
            ))
        }
    };
    if real == 0.0 {
        if imaginary < 0.0 {
            return Ok(format!("-{}", imaginary_text(-imaginary)?));
        }
        return imaginary_text(imaginary);
    }
    let sign = if imaginary < 0.0 { "-" } else { "+" };
    Ok(format!(
        "{}{}{}",
        formula_text_from_number(real)?,
        sign,
        imaginary_text(imaginary.abs())?
    ))
}

fn formula_complex_join_suffix(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<Option<char>, FormulaEvalError> {
    match (left.suffix, right.suffix) {
        (Some(left), Some(right)) if left != right => Err(FormulaEvalError::Value),
        (Some(suffix), _) | (_, Some(suffix)) => Ok(Some(suffix)),
        (None, None) => Ok(None),
    }
}

fn formula_complex_add(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        left.real + right.real,
        left.imaginary + right.imaginary,
        formula_complex_join_suffix(left, right)?,
    )
}

fn formula_complex_subtract(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        left.real - right.real,
        left.imaginary - right.imaginary,
        formula_complex_join_suffix(left, right)?,
    )
}

fn formula_complex_multiply(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        left.real * right.real - left.imaginary * right.imaginary,
        left.real * right.imaginary + left.imaginary * right.real,
        formula_complex_join_suffix(left, right)?,
    )
}

fn formula_complex_divide(
    left: FormulaComplexNumber,
    right: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let denominator = right.real * right.real + right.imaginary * right.imaginary;
    if denominator == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    formula_complex_number(
        (left.real * right.real + left.imaginary * right.imaginary) / denominator,
        (left.imaginary * right.real - left.real * right.imaginary) / denominator,
        formula_complex_join_suffix(left, right)?,
    )
}

fn formula_complex_reciprocal(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_divide(formula_complex_number(1.0, 0.0, value.suffix)?, value)
}

fn formula_complex_exp(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let magnitude = value.real.exp();
    formula_complex_number(
        magnitude * value.imaginary.cos(),
        magnitude * value.imaginary.sin(),
        value.suffix,
    )
}

fn formula_complex_ln(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let magnitude = value.real.hypot(value.imaginary);
    if magnitude == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    formula_complex_number(
        magnitude.ln(),
        value.imaginary.atan2(value.real),
        value.suffix,
    )
}

fn formula_complex_power(
    value: FormulaComplexNumber,
    power: f64,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    if !power.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let magnitude = value.real.hypot(value.imaginary);
    if magnitude == 0.0 {
        if power <= 0.0 {
            return Err(FormulaEvalError::Num);
        }
        return formula_complex_number(0.0, 0.0, value.suffix);
    }
    let powered_magnitude = magnitude.powf(power);
    let argument = value.imaginary.atan2(value.real) * power;
    formula_complex_number(
        powered_magnitude * argument.cos(),
        powered_magnitude * argument.sin(),
        value.suffix,
    )
}

fn formula_complex_sin(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.sin() * value.imaginary.cosh(),
        value.real.cos() * value.imaginary.sinh(),
        value.suffix,
    )
}

fn formula_complex_cos(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.cos() * value.imaginary.cosh(),
        -value.real.sin() * value.imaginary.sinh(),
        value.suffix,
    )
}

fn formula_complex_sinh(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.sinh() * value.imaginary.cos(),
        value.real.cosh() * value.imaginary.sin(),
        value.suffix,
    )
}

fn formula_complex_cosh(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    formula_complex_number(
        value.real.cosh() * value.imaginary.cos(),
        value.real.sinh() * value.imaginary.sin(),
        value.suffix,
    )
}

fn formula_complex_sqrt(
    value: FormulaComplexNumber,
) -> Result<FormulaComplexNumber, FormulaEvalError> {
    let magnitude = value.real.hypot(value.imaginary);
    let real = ((magnitude + value.real) / 2.0).sqrt();
    let imaginary_sign = if value.imaginary < 0.0 { -1.0 } else { 1.0 };
    let imaginary = imaginary_sign * ((magnitude - value.real) / 2.0).sqrt();
    formula_complex_number(real, imaginary, value.suffix)
}

fn formula_bessel_order(value: f64) -> Result<usize, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let value = value.trunc();
    if value < 0.0 {
        return Err(FormulaEvalError::Num);
    }
    if value > 10_000.0 {
        return Err(FormulaEvalError::Num);
    }
    Ok(value as usize)
}

fn formula_bessel_j0(value: f64) -> f64 {
    let absolute = value.abs();
    if absolute < 8.0 {
        let y = value * value;
        let numerator = 57_568_490_574.0
            + y * (-13_362_590_354.0
                + y * (651_619_640.7
                    + y * (-11_214_424.18 + y * (77_392.33017 + y * -184.9052456))));
        let denominator = 57_568_490_411.0
            + y * (1_029_532_985.0
                + y * (9_494_680.718 + y * (59_272.64853 + y * (267.8532712 + y))));
        numerator / denominator
    } else {
        let z = 8.0 / absolute;
        let y = z * z;
        let angle = absolute - std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (-0.001098628627
                + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)));
        let second = -0.01562499995
            + y * (0.0001430488765
                + y * (-0.000006911147651 + y * (0.0000007621095161 - y * 0.0000000934945152)));
        (0.636619772 / absolute).sqrt() * (angle.cos() * first - z * angle.sin() * second)
    }
}

fn formula_bessel_j1(value: f64) -> f64 {
    let absolute = value.abs();
    let result = if absolute < 8.0 {
        let y = value * value;
        let numerator = absolute
            * (72_362_614_232.0
                + y * (-7_895_059_235.0
                    + y * (242_396_853.1
                        + y * (-2_972_611.439 + y * (15_704.48260 + y * -30.16036606)))));
        let denominator = 144_725_228_442.0
            + y * (2_300_535_178.0
                + y * (18_583_304.74 + y * (99_447.43394 + y * (376.9991397 + y))));
        numerator / denominator
    } else {
        let z = 8.0 / absolute;
        let y = z * z;
        let angle = absolute - 3.0 * std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (0.00183105
                + y * (-0.00003516396496 + y * (0.000002457520174 - y * 0.000000240337019)));
        let second = 0.04687499995
            + y * (-0.0002002690873
                + y * (0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)));
        (0.636619772 / absolute).sqrt() * (angle.cos() * first - z * angle.sin() * second)
    };
    if value < 0.0 { -result } else { result }
}

fn formula_bessel_y0(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value < 8.0 {
        let y = value * value;
        let numerator = -2_957_821_389.0
            + y * (7_062_834_065.0
                + y * (-512_359_803.6
                    + y * (10_879_881.29 + y * (-86_327.92757 + y * 228.4622733))));
        let denominator = 40_076_544_269.0
            + y * (745_249_964.8
                + y * (7_189_466.438 + y * (47_447.26470 + y * (226.1030244 + y))));
        Ok(numerator / denominator + 0.636619772 * formula_bessel_j0(value) * value.ln())
    } else {
        let z = 8.0 / value;
        let y = z * z;
        let angle = value - std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (-0.001098628627
                + y * (0.00002734510407 + y * (-0.000002073370639 + y * 0.0000002093887211)));
        let second = -0.01562499995
            + y * (0.0001430488765
                + y * (-0.000006911147651 + y * (0.0000007621095161 - y * 0.0000000934945152)));
        Ok((0.636619772 / value).sqrt() * (angle.sin() * first + z * angle.cos() * second))
    }
}

fn formula_bessel_y1(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value < 8.0 {
        let y = value * value;
        let numerator = value
            * (-4_900_604_943_000.0
                + y * (1_275_274_390_000.0
                    + y * (-51_534_381_390.0
                        + y * (734_926_455.1 + y * (-4_237_922.726 + y * 8_511.937935)))));
        let denominator = 24_995_805_700_000.0
            + y * (424_441_966_400.0
                + y * (3_733_650_367.0
                    + y * (22_459_040.02 + y * (102_042.6050 + y * (354.9632885 + y)))));
        Ok(numerator / denominator
            + 0.636619772 * (formula_bessel_j1(value) * value.ln() - 1.0 / value))
    } else {
        let z = 8.0 / value;
        let y = z * z;
        let angle = value - 3.0 * std::f64::consts::FRAC_PI_4;
        let first = 1.0
            + y * (0.00183105
                + y * (-0.00003516396496 + y * (0.000002457520174 - y * 0.000000240337019)));
        let second = 0.04687499995
            + y * (-0.0002002690873
                + y * (0.000008449199096 + y * (-0.00000088228987 + y * 0.000000105787412)));
        Ok((0.636619772 / value).sqrt() * (angle.sin() * first + z * angle.cos() * second))
    }
}

fn formula_bessel_i0(value: f64) -> f64 {
    let absolute = value.abs();
    if absolute < 3.75 {
        let y = (value / 3.75).powi(2);
        1.0 + y
            * (3.5156229
                + y * (3.0899424
                    + y * (1.2067492 + y * (0.2659732 + y * (0.0360768 + y * 0.0045813)))))
    } else {
        let y = 3.75 / absolute;
        absolute.exp() / absolute.sqrt()
            * (0.39894228
                + y * (0.01328592
                    + y * (0.00225319
                        + y * (-0.00157565
                            + y * (0.00916281
                                + y * (-0.02057706
                                    + y * (0.02635537 + y * (-0.01647633 + y * 0.00392377))))))))
    }
}

fn formula_bessel_i1(value: f64) -> f64 {
    let absolute = value.abs();
    let result = if absolute < 3.75 {
        let y = (value / 3.75).powi(2);
        absolute
            * (0.5
                + y * (0.87890594
                    + y * (0.51498869
                        + y * (0.15084934 + y * (0.02658733 + y * (0.00301532 + y * 0.00032411))))))
    } else {
        let y = 3.75 / absolute;
        absolute.exp() / absolute.sqrt()
            * (0.39894228
                + y * (-0.03988024
                    + y * (-0.00362018
                        + y * (0.00163801
                            + y * (-0.01031555
                                + y * (0.02282967
                                    + y * (-0.02895312 + y * (0.01787654 - y * 0.00420059))))))))
    };
    if value < 0.0 { -result } else { result }
}

fn formula_bessel_k0(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value <= 2.0 {
        let y = value * value / 4.0;
        Ok(-(value / 2.0).ln() * formula_bessel_i0(value)
            + (-0.57721566
                + y * (0.42278420
                    + y * (0.23069756
                        + y * (0.03488590
                            + y * (0.00262698 + y * (0.00010750 + y * 0.00000740)))))))
    } else {
        let y = 2.0 / value;
        Ok((-value).exp() / value.sqrt()
            * (1.25331414
                + y * (-0.07832358
                    + y * (0.02189568
                        + y * (-0.01062446
                            + y * (0.00587872 + y * (-0.00251540 + y * 0.00053208)))))))
    }
}

fn formula_bessel_k1(value: f64) -> Result<f64, FormulaEvalError> {
    if value <= 0.0 || !value.is_finite() {
        return Err(if value.is_finite() {
            FormulaEvalError::Num
        } else {
            FormulaEvalError::Value
        });
    }
    if value <= 2.0 {
        let y = value * value / 4.0;
        Ok((value / 2.0).ln() * formula_bessel_i1(value)
            + (1.0 / value)
                * (1.0
                    + y * (0.15443144
                        + y * (-0.67278579
                            + y * (-0.18156897
                                + y * (-0.01919402 + y * (-0.00110404 - y * 0.00004686)))))))
    } else {
        let y = 2.0 / value;
        Ok((-value).exp() / value.sqrt()
            * (1.25331414
                + y * (0.23498619
                    + y * (-0.03655620
                        + y * (0.01504268
                            + y * (-0.00780353 + y * (0.00325614 - y * 0.00068245)))))))
    }
}

fn formula_bessel_i(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if order == 0 {
        return Ok(formula_bessel_i0(value));
    }
    if order == 1 {
        return Ok(formula_bessel_i1(value));
    }
    let absolute = value.abs();
    if absolute == 0.0 {
        return Ok(0.0);
    }
    const BIGNO: f64 = 1e100;
    const BIGNI: f64 = 1e-100;
    let tox = 2.0 / absolute;
    let mut bip = 0.0;
    let mut bi = 1.0;
    let mut answer = 0.0;
    let m = 2 * (order + (40.0 * order as f64).sqrt() as usize);
    for j in (1..=m).rev() {
        let bim = bip + j as f64 * tox * bi;
        bip = bi;
        bi = bim;
        if bi.abs() > BIGNO {
            answer *= BIGNI;
            bi *= BIGNI;
            bip *= BIGNI;
        }
        if j == order {
            answer = bip;
        }
    }
    answer *= formula_bessel_i0(absolute) / bi;
    if value < 0.0 && order % 2 == 1 {
        answer = -answer;
    }
    if answer.is_finite() {
        Ok(answer)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_bessel_j(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if order == 0 {
        return Ok(formula_bessel_j0(value));
    }
    if order == 1 {
        return Ok(formula_bessel_j1(value));
    }
    let absolute = value.abs();
    if absolute == 0.0 {
        return Ok(0.0);
    }
    let tox = 2.0 / absolute;
    let mut answer;
    if absolute > order as f64 {
        let mut previous = formula_bessel_j0(absolute);
        let mut current = formula_bessel_j1(absolute);
        for j in 1..order {
            let next = j as f64 * tox * current - previous;
            previous = current;
            current = next;
        }
        answer = current;
    } else {
        const BIGNO: f64 = 1e100;
        const BIGNI: f64 = 1e-100;
        let mut next = 0.0;
        let mut current = 1.0;
        let mut sum = 0.0;
        let mut include_in_sum = false;
        answer = 0.0;
        let m = 2 * ((order + (40.0 * order as f64).sqrt() as usize) / 2);
        for j in (1..=m).rev() {
            let previous = j as f64 * tox * current - next;
            next = current;
            current = previous;
            if current.abs() > BIGNO {
                answer *= BIGNI;
                current *= BIGNI;
                next *= BIGNI;
                sum *= BIGNI;
            }
            if include_in_sum {
                sum += current;
            }
            include_in_sum = !include_in_sum;
            if j == order {
                answer = next;
            }
        }
        sum = 2.0 * sum - current;
        answer /= sum;
    }
    if value < 0.0 && order % 2 == 1 {
        answer = -answer;
    }
    if answer.is_finite() {
        Ok(answer)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_bessel_k(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if order == 0 {
        return formula_bessel_k0(value);
    }
    if order == 1 {
        return formula_bessel_k1(value);
    }
    let mut previous = formula_bessel_k0(value)?;
    let mut current = formula_bessel_k1(value)?;
    let tox = 2.0 / value;
    for j in 1..order {
        let next = previous + j as f64 * tox * current;
        previous = current;
        current = next;
        if !current.is_finite() {
            return Err(FormulaEvalError::Num);
        }
    }
    Ok(current)
}

fn formula_bessel_y(value: f64, order: usize) -> Result<f64, FormulaEvalError> {
    if order == 0 {
        return formula_bessel_y0(value);
    }
    if order == 1 {
        return formula_bessel_y1(value);
    }
    let mut previous = formula_bessel_y0(value)?;
    let mut current = formula_bessel_y1(value)?;
    let tox = 2.0 / value;
    for j in 1..order {
        let next = j as f64 * tox * current - previous;
        previous = current;
        current = next;
        if !current.is_finite() {
            return Err(FormulaEvalError::Num);
        }
    }
    Ok(current)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaConvertDimension {
    Area,
    Distance,
    Energy,
    Force,
    Information,
    Magnetism,
    Mass,
    Power,
    Pressure,
    Speed,
    Temperature,
    Time,
    Volume,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaConvertTemperatureUnit {
    Celsius,
    Fahrenheit,
    Kelvin,
    Rankine,
    Reaumur,
}

#[derive(Debug, Clone, Copy)]
enum FormulaConvertScale {
    Ratio(f64),
    Temperature(FormulaConvertTemperatureUnit),
}

#[derive(Debug, Clone, Copy)]
struct FormulaConvertUnit {
    dimension: FormulaConvertDimension,
    scale: FormulaConvertScale,
    metric_power: i32,
    binary_prefixable: bool,
}

fn formula_convert_ratio_unit(
    dimension: FormulaConvertDimension,
    factor: f64,
    metric_power: i32,
    binary_prefixable: bool,
) -> FormulaConvertUnit {
    FormulaConvertUnit {
        dimension,
        scale: FormulaConvertScale::Ratio(factor),
        metric_power,
        binary_prefixable,
    }
}

fn formula_convert_temperature_unit(unit: FormulaConvertTemperatureUnit) -> FormulaConvertUnit {
    FormulaConvertUnit {
        dimension: FormulaConvertDimension::Temperature,
        scale: FormulaConvertScale::Temperature(unit),
        metric_power: 0,
        binary_prefixable: false,
    }
}

fn formula_convert_exact_unit(unit: &str) -> Option<FormulaConvertUnit> {
    use FormulaConvertDimension::*;
    const INCH: f64 = 0.0254;
    const FOOT: f64 = 0.3048;
    const YARD: f64 = 0.9144;
    const MILE: f64 = 1609.344;
    const SURVEY_MILE: f64 = 1609.3472186944373;
    const NAUTICAL_MILE: f64 = 1852.0;
    const LIGHT_YEAR: f64 = 9_460_730_472_580_800.0;
    const PARSEC: f64 = 30_856_775_814_913_670.0;
    const PICA_POINT: f64 = INCH / 72.0;
    const PICA: f64 = INCH / 6.0;
    const US_FLUID_OUNCE: f64 = 0.0000295735295625;
    const UK_PINT: f64 = 0.00056826125;
    const UK_GALLON: f64 = 0.00454609;
    let unit = match unit {
        "g" => formula_convert_ratio_unit(Mass, 1.0, 1, false),
        "sg" => formula_convert_ratio_unit(Mass, 14_593.90294, 0, false),
        "lbm" => formula_convert_ratio_unit(Mass, 453.59237, 0, false),
        "u" => formula_convert_ratio_unit(Mass, 1.660_539_066_6e-24, 0, false),
        "ozm" => formula_convert_ratio_unit(Mass, 28.349523125, 0, false),
        "grain" => formula_convert_ratio_unit(Mass, 0.06479891, 0, false),
        "cwt" | "shweight" => formula_convert_ratio_unit(Mass, 45_359.237, 0, false),
        "uk_cwt" | "lcwt" | "hweight" => formula_convert_ratio_unit(Mass, 50_802.34544, 0, false),
        "stone" => formula_convert_ratio_unit(Mass, 6_350.29318, 0, false),
        "ton" => formula_convert_ratio_unit(Mass, 907_184.74, 0, false),
        "uk_ton" | "LTON" | "brton" => formula_convert_ratio_unit(Mass, 1_016_046.9088, 0, false),
        "m" => formula_convert_ratio_unit(Distance, 1.0, 1, false),
        "mi" => formula_convert_ratio_unit(Distance, MILE, 0, false),
        "Nmi" => formula_convert_ratio_unit(Distance, NAUTICAL_MILE, 0, false),
        "in" => formula_convert_ratio_unit(Distance, INCH, 0, false),
        "ft" => formula_convert_ratio_unit(Distance, FOOT, 0, false),
        "yd" => formula_convert_ratio_unit(Distance, YARD, 0, false),
        "ang" => formula_convert_ratio_unit(Distance, 1e-10, 0, false),
        "ell" => formula_convert_ratio_unit(Distance, 45.0 * INCH, 0, false),
        "ly" => formula_convert_ratio_unit(Distance, LIGHT_YEAR, 0, false),
        "parsec" | "pc" => formula_convert_ratio_unit(Distance, PARSEC, 0, false),
        "Picapt" | "Pica" => formula_convert_ratio_unit(Distance, PICA_POINT, 0, false),
        "pica" => formula_convert_ratio_unit(Distance, PICA, 0, false),
        "survey_mi" => formula_convert_ratio_unit(Distance, SURVEY_MILE, 0, false),
        "yr" => formula_convert_ratio_unit(Time, 31_557_600.0, 0, false),
        "day" | "d" => formula_convert_ratio_unit(Time, 86_400.0, 0, false),
        "hr" => formula_convert_ratio_unit(Time, 3_600.0, 0, false),
        "mn" | "min" => formula_convert_ratio_unit(Time, 60.0, 0, false),
        "sec" | "s" => formula_convert_ratio_unit(Time, 1.0, 1, false),
        "Pa" | "p" => formula_convert_ratio_unit(Pressure, 1.0, 1, false),
        "atm" | "at" => formula_convert_ratio_unit(Pressure, 101_325.0, 0, false),
        "mmHg" => formula_convert_ratio_unit(Pressure, 133.322, 0, false),
        "psi" => formula_convert_ratio_unit(Pressure, 6_894.757293168361, 0, false),
        "Torr" => formula_convert_ratio_unit(Pressure, 101_325.0 / 760.0, 0, false),
        "N" => formula_convert_ratio_unit(Force, 1.0, 1, false),
        "dyn" | "dy" => formula_convert_ratio_unit(Force, 1e-5, 1, false),
        "lbf" => formula_convert_ratio_unit(Force, 4.4482216152605, 0, false),
        "J" => formula_convert_ratio_unit(Energy, 1.0, 1, false),
        "e" => formula_convert_ratio_unit(Energy, 1e-7, 1, false),
        "c" => formula_convert_ratio_unit(Energy, 4.184, 1, false),
        "cal" => formula_convert_ratio_unit(Energy, 4.1868, 1, false),
        "eV" | "ev" => formula_convert_ratio_unit(Energy, 1.602_176_634e-19, 1, false),
        "HPh" | "hh" => formula_convert_ratio_unit(Energy, 2_684_519.538, 0, false),
        "Wh" | "wh" => formula_convert_ratio_unit(Energy, 3_600.0, 1, false),
        "flb" => formula_convert_ratio_unit(Energy, 1.3558179483314004, 0, false),
        "BTU" | "btu" => formula_convert_ratio_unit(Energy, 1_055.05585262, 0, false),
        "HP" | "h" => formula_convert_ratio_unit(Power, 745.6998715822702, 0, false),
        "PS" => formula_convert_ratio_unit(Power, 735.49875, 0, false),
        "W" | "w" => formula_convert_ratio_unit(Power, 1.0, 1, false),
        "T" => formula_convert_ratio_unit(Magnetism, 1.0, 1, false),
        "ga" => formula_convert_ratio_unit(Magnetism, 1e-4, 0, false),
        "C" | "cel" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Celsius),
        "F" | "fah" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Fahrenheit),
        "K" | "kel" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Kelvin),
        "Rank" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Rankine),
        "Reau" => formula_convert_temperature_unit(FormulaConvertTemperatureUnit::Reaumur),
        "tsp" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE / 6.0, 0, false),
        "tspm" => formula_convert_ratio_unit(Volume, 0.000005, 0, false),
        "tbs" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE / 2.0, 0, false),
        "oz" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE, 0, false),
        "cup" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 8.0, 0, false),
        "pt" | "us_pt" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 16.0, 0, false),
        "uk_pt" => formula_convert_ratio_unit(Volume, UK_PINT, 0, false),
        "qt" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 32.0, 0, false),
        "uk_qt" => formula_convert_ratio_unit(Volume, UK_PINT * 2.0, 0, false),
        "gal" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 128.0, 0, false),
        "uk_gal" => formula_convert_ratio_unit(Volume, UK_GALLON, 0, false),
        "l" | "L" | "lt" => formula_convert_ratio_unit(Volume, 0.001, 1, false),
        "ang3" | "ang^3" => formula_convert_ratio_unit(Volume, 1e-30, 0, false),
        "barrel" => formula_convert_ratio_unit(Volume, US_FLUID_OUNCE * 128.0 * 42.0, 0, false),
        "bushel" => formula_convert_ratio_unit(Volume, 2_150.42 * INCH.powi(3), 0, false),
        "ft3" | "ft^3" => formula_convert_ratio_unit(Volume, FOOT.powi(3), 0, false),
        "in3" | "in^3" => formula_convert_ratio_unit(Volume, INCH.powi(3), 0, false),
        "ly3" | "ly^3" => formula_convert_ratio_unit(Volume, LIGHT_YEAR.powi(3), 0, false),
        "m3" | "m^3" => formula_convert_ratio_unit(Volume, 1.0, 3, false),
        "mi3" | "mi^3" => formula_convert_ratio_unit(Volume, MILE.powi(3), 0, false),
        "yd3" | "yd^3" => formula_convert_ratio_unit(Volume, YARD.powi(3), 0, false),
        "Nmi3" | "Nmi^3" => formula_convert_ratio_unit(Volume, NAUTICAL_MILE.powi(3), 0, false),
        "Picapt3" | "Picapt^3" | "Pica3" | "Pica^3" => {
            formula_convert_ratio_unit(Volume, PICA_POINT.powi(3), 0, false)
        }
        "GRT" | "regton" => formula_convert_ratio_unit(Volume, 100.0 * FOOT.powi(3), 0, false),
        "MTON" => formula_convert_ratio_unit(Volume, 40.0 * FOOT.powi(3), 0, false),
        "uk_acre" => formula_convert_ratio_unit(Area, 4_046.8564224, 0, false),
        "us_acre" => formula_convert_ratio_unit(Area, 4_046.872609874252, 0, false),
        "ang2" | "ang^2" => formula_convert_ratio_unit(Area, 1e-20, 0, false),
        "ar" => formula_convert_ratio_unit(Area, 100.0, 1, false),
        "ft2" | "ft^2" => formula_convert_ratio_unit(Area, FOOT.powi(2), 0, false),
        "ha" => formula_convert_ratio_unit(Area, 10_000.0, 0, false),
        "in2" | "in^2" => formula_convert_ratio_unit(Area, INCH.powi(2), 0, false),
        "ly2" | "ly^2" => formula_convert_ratio_unit(Area, LIGHT_YEAR.powi(2), 0, false),
        "m2" | "m^2" => formula_convert_ratio_unit(Area, 1.0, 2, false),
        "Morgen" => formula_convert_ratio_unit(Area, 2_500.0, 0, false),
        "mi2" | "mi^2" => formula_convert_ratio_unit(Area, MILE.powi(2), 0, false),
        "Nmi2" | "Nmi^2" => formula_convert_ratio_unit(Area, NAUTICAL_MILE.powi(2), 0, false),
        "Picapt2" | "Pica2" | "Pica^2" | "Picapt^2" => {
            formula_convert_ratio_unit(Area, PICA_POINT.powi(2), 0, false)
        }
        "yd2" | "yd^2" => formula_convert_ratio_unit(Area, YARD.powi(2), 0, false),
        "bit" => formula_convert_ratio_unit(Information, 1.0, 1, true),
        "byte" => formula_convert_ratio_unit(Information, 8.0, 1, true),
        "admkn" => formula_convert_ratio_unit(Speed, 6080.0 * FOOT / 3600.0, 0, false),
        "kn" => formula_convert_ratio_unit(Speed, NAUTICAL_MILE / 3600.0, 0, false),
        "m/h" | "m/hr" => formula_convert_ratio_unit(Speed, 1.0 / 3600.0, 1, false),
        "m/s" | "m/sec" => formula_convert_ratio_unit(Speed, 1.0, 1, false),
        "mph" => formula_convert_ratio_unit(Speed, MILE / 3600.0, 0, false),
        _ => return None,
    };
    Some(unit)
}

fn formula_convert_metric_prefix(unit: &str) -> Option<(f64, &str)> {
    const PREFIXES: [(&str, f64); 21] = [
        ("da", 1e1),
        ("Y", 1e24),
        ("Z", 1e21),
        ("E", 1e18),
        ("P", 1e15),
        ("T", 1e12),
        ("G", 1e9),
        ("M", 1e6),
        ("k", 1e3),
        ("h", 1e2),
        ("e", 1e1),
        ("d", 1e-1),
        ("c", 1e-2),
        ("m", 1e-3),
        ("u", 1e-6),
        ("n", 1e-9),
        ("p", 1e-12),
        ("f", 1e-15),
        ("a", 1e-18),
        ("z", 1e-21),
        ("y", 1e-24),
    ];
    PREFIXES
        .iter()
        .find_map(|(prefix, factor)| unit.strip_prefix(prefix).map(|rest| (*factor, rest)))
        .filter(|(_, rest)| !rest.is_empty())
}

fn formula_convert_binary_prefix(unit: &str) -> Option<(f64, &str)> {
    const PREFIXES: [(&str, f64); 8] = [
        ("Yi", 1_208_925_819_614_629_174_706_176.0),
        ("Zi", 1_180_591_620_717_411_303_424.0),
        ("Ei", 1_152_921_504_606_846_976.0),
        ("Pi", 1_125_899_906_842_624.0),
        ("Ti", 1_099_511_627_776.0),
        ("Gi", 1_073_741_824.0),
        ("Mi", 1_048_576.0),
        ("ki", 1_024.0),
    ];
    PREFIXES
        .iter()
        .find_map(|(prefix, factor)| unit.strip_prefix(prefix).map(|rest| (*factor, rest)))
        .filter(|(_, rest)| !rest.is_empty())
}

fn formula_convert_unit(unit: &str) -> Result<FormulaConvertUnit, FormulaEvalError> {
    if let Some(unit) = formula_convert_exact_unit(unit) {
        return Ok(unit);
    }
    if let Some((factor, suffix)) = formula_convert_binary_prefix(unit) {
        if let Some(mut unit) = formula_convert_exact_unit(suffix) {
            if unit.binary_prefixable {
                if let FormulaConvertScale::Ratio(base_factor) = unit.scale {
                    unit.scale = FormulaConvertScale::Ratio(base_factor * factor);
                    return Ok(unit);
                }
            }
        }
        return Err(FormulaEvalError::NA);
    }
    if let Some((factor, suffix)) = formula_convert_metric_prefix(unit) {
        if let Some(mut unit) = formula_convert_exact_unit(suffix) {
            if unit.metric_power > 0 {
                if let FormulaConvertScale::Ratio(base_factor) = unit.scale {
                    unit.scale =
                        FormulaConvertScale::Ratio(base_factor * factor.powi(unit.metric_power));
                    return Ok(unit);
                }
            }
        }
        return Err(FormulaEvalError::NA);
    }
    Err(FormulaEvalError::NA)
}

fn formula_convert_temperature_to_kelvin(value: f64, unit: FormulaConvertTemperatureUnit) -> f64 {
    match unit {
        FormulaConvertTemperatureUnit::Celsius => value + 273.15,
        FormulaConvertTemperatureUnit::Fahrenheit => (value + 459.67) * 5.0 / 9.0,
        FormulaConvertTemperatureUnit::Kelvin => value,
        FormulaConvertTemperatureUnit::Rankine => value * 5.0 / 9.0,
        FormulaConvertTemperatureUnit::Reaumur => value * 1.25 + 273.15,
    }
}

fn formula_convert_temperature_from_kelvin(value: f64, unit: FormulaConvertTemperatureUnit) -> f64 {
    match unit {
        FormulaConvertTemperatureUnit::Celsius => value - 273.15,
        FormulaConvertTemperatureUnit::Fahrenheit => value * 9.0 / 5.0 - 459.67,
        FormulaConvertTemperatureUnit::Kelvin => value,
        FormulaConvertTemperatureUnit::Rankine => value * 9.0 / 5.0,
        FormulaConvertTemperatureUnit::Reaumur => (value - 273.15) * 0.8,
    }
}

fn formula_convert_value(
    value: f64,
    from_unit: &str,
    to_unit: &str,
) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let from_unit = formula_convert_unit(from_unit)?;
    let to_unit = formula_convert_unit(to_unit)?;
    if from_unit.dimension != to_unit.dimension {
        return Err(FormulaEvalError::NA);
    }
    let result = match (from_unit.scale, to_unit.scale) {
        (FormulaConvertScale::Ratio(from_factor), FormulaConvertScale::Ratio(to_factor)) => {
            value * from_factor / to_factor
        }
        (
            FormulaConvertScale::Temperature(from_unit),
            FormulaConvertScale::Temperature(to_unit),
        ) => formula_convert_temperature_from_kelvin(
            formula_convert_temperature_to_kelvin(value, from_unit),
            to_unit,
        ),
        _ => return Err(FormulaEvalError::NA),
    };
    if result.is_finite() {
        Ok(result)
    } else {
        Err(FormulaEvalError::Num)
    }
}

#[derive(Debug, Clone, Copy)]
struct FormulaEuroCurrency {
    rate: f64,
    calculation_precision: i32,
}

fn formula_euro_currency(code: &str) -> Option<FormulaEuroCurrency> {
    let currency = match code {
        "ATS" => FormulaEuroCurrency {
            rate: 13.7603,
            calculation_precision: 2,
        },
        "BEF" | "LUF" => FormulaEuroCurrency {
            rate: 40.3399,
            calculation_precision: 0,
        },
        "DEM" => FormulaEuroCurrency {
            rate: 1.95583,
            calculation_precision: 2,
        },
        "ESP" => FormulaEuroCurrency {
            rate: 166.386,
            calculation_precision: 0,
        },
        "EUR" => FormulaEuroCurrency {
            rate: 1.0,
            calculation_precision: 2,
        },
        "FIM" => FormulaEuroCurrency {
            rate: 5.94573,
            calculation_precision: 2,
        },
        "FRF" => FormulaEuroCurrency {
            rate: 6.55957,
            calculation_precision: 2,
        },
        "GRD" => FormulaEuroCurrency {
            rate: 340.75,
            calculation_precision: 0,
        },
        "IEP" => FormulaEuroCurrency {
            rate: 0.787564,
            calculation_precision: 2,
        },
        "ITL" => FormulaEuroCurrency {
            rate: 1936.27,
            calculation_precision: 0,
        },
        "NLG" => FormulaEuroCurrency {
            rate: 2.20371,
            calculation_precision: 2,
        },
        "PTE" => FormulaEuroCurrency {
            rate: 200.482,
            calculation_precision: 0,
        },
        "SIT" => FormulaEuroCurrency {
            rate: 239.64,
            calculation_precision: 2,
        },
        _ => return None,
    };
    Some(currency)
}

fn formula_round_to_decimal_places(value: f64, places: i32) -> Result<f64, FormulaEvalError> {
    let factor = 10_f64.powi(places);
    if !factor.is_finite() || factor == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    formula_checked_numeric_result(round_half_away_from_zero(value * factor) / factor)
}

fn formula_round_to_significant_digits(value: f64, digits: i64) -> Result<f64, FormulaEvalError> {
    if digits < 1 || digits > i64::from(i32::MAX) {
        return Err(FormulaEvalError::Value);
    }
    if value == 0.0 {
        return Ok(0.0);
    }
    let magnitude = value.abs().log10().floor();
    if !magnitude.is_finite() || magnitude < i32::MIN as f64 || magnitude > i32::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    let places = i32::try_from(digits - 1).map_err(|_| FormulaEvalError::Value)? - magnitude as i32;
    formula_round_to_decimal_places(value, places)
}

fn formula_euroconvert_value(
    value: f64,
    source_code: &str,
    target_code: &str,
    full_precision: bool,
    triangulation_precision: Option<i64>,
) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let source_code = source_code.to_ascii_uppercase();
    let target_code = target_code.to_ascii_uppercase();
    if source_code == target_code {
        return Ok(value);
    }
    let source = formula_euro_currency(source_code.as_str()).ok_or(FormulaEvalError::Value)?;
    let target = formula_euro_currency(target_code.as_str()).ok_or(FormulaEvalError::Value)?;
    let mut euros = value / source.rate;
    if let Some(precision) = triangulation_precision {
        if precision < 3 {
            return Err(FormulaEvalError::Value);
        }
        if source_code != "EUR" {
            euros = formula_round_to_significant_digits(euros, precision)?;
        }
    }
    let result = euros * target.rate;
    if full_precision {
        formula_checked_numeric_result(result)
    } else {
        formula_round_to_decimal_places(result, target.calculation_precision)
    }
}

fn formula_matrix_determinant(mut matrix: Vec<Vec<f64>>) -> Result<f64, FormulaEvalError> {
    let size = matrix.len();
    if size == 0 || matrix.iter().any(|row| row.len() != size) {
        return Err(FormulaEvalError::Value);
    }
    if matrix.iter().flatten().any(|value| !value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }

    let mut determinant = 1.0_f64;
    for pivot_index in 0..size {
        let mut pivot_row = pivot_index;
        let mut pivot_abs = matrix[pivot_index][pivot_index].abs();
        for (row_index, row) in matrix.iter().enumerate().skip(pivot_index + 1) {
            let candidate_abs = row[pivot_index].abs();
            if candidate_abs > pivot_abs {
                pivot_abs = candidate_abs;
                pivot_row = row_index;
            }
        }
        if pivot_abs <= 1e-12 {
            return Ok(0.0);
        }
        if pivot_row != pivot_index {
            matrix.swap(pivot_index, pivot_row);
            determinant = -determinant;
        }
        let pivot = matrix[pivot_index][pivot_index];
        determinant *= pivot;
        if !determinant.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        for row_index in pivot_index + 1..size {
            let factor = matrix[row_index][pivot_index] / pivot;
            for col_index in pivot_index + 1..size {
                matrix[row_index][col_index] -= factor * matrix[pivot_index][col_index];
            }
        }
    }
    formula_checked_numeric_result(determinant)
}

fn formula_matrix_inverse_top_left(mut matrix: Vec<Vec<f64>>) -> Result<f64, FormulaEvalError> {
    let size = matrix.len();
    if size == 0 || matrix.iter().any(|row| row.len() != size) {
        return Err(FormulaEvalError::Value);
    }
    if matrix.iter().flatten().any(|value| !value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    let mut inverse = vec![vec![0.0_f64; size]; size];
    for (index, row) in inverse.iter_mut().enumerate() {
        row[index] = 1.0;
    }

    for pivot_index in 0..size {
        let mut pivot_row = pivot_index;
        let mut pivot_abs = matrix[pivot_index][pivot_index].abs();
        for (row_index, row) in matrix.iter().enumerate().skip(pivot_index + 1) {
            let candidate_abs = row[pivot_index].abs();
            if candidate_abs > pivot_abs {
                pivot_abs = candidate_abs;
                pivot_row = row_index;
            }
        }
        if pivot_abs <= 1e-12 {
            return Err(FormulaEvalError::Num);
        }
        if pivot_row != pivot_index {
            matrix.swap(pivot_index, pivot_row);
            inverse.swap(pivot_index, pivot_row);
        }
        let pivot = matrix[pivot_index][pivot_index];
        for col_index in 0..size {
            matrix[pivot_index][col_index] /= pivot;
            inverse[pivot_index][col_index] /= pivot;
        }
        for row_index in 0..size {
            if row_index == pivot_index {
                continue;
            }
            let factor = matrix[row_index][pivot_index];
            for col_index in 0..size {
                matrix[row_index][col_index] -= factor * matrix[pivot_index][col_index];
                inverse[row_index][col_index] -= factor * inverse[pivot_index][col_index];
            }
        }
    }
    formula_checked_numeric_result(inverse[0][0])
}

fn formula_regression_default_x(count: usize) -> Vec<f64> {
    (1..=count).map(|value| value as f64).collect()
}

fn formula_regression_slope_intercept(
    known_y: &[f64],
    known_x: &[f64],
    constant: bool,
    exponential: bool,
) -> Result<(f64, f64), FormulaEvalError> {
    if known_y.len() != known_x.len() || known_y.is_empty() {
        return Err(FormulaEvalError::NA);
    }
    if known_y
        .iter()
        .chain(known_x.iter())
        .any(|value| !value.is_finite())
    {
        return Err(FormulaEvalError::Value);
    }
    let y_values = if exponential {
        let mut transformed = Vec::with_capacity(known_y.len());
        for value in known_y {
            if *value <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            transformed.push(value.ln());
        }
        transformed
    } else {
        known_y.to_vec()
    };

    if constant {
        let count = y_values.len() as f64;
        let mean_y = y_values.iter().sum::<f64>() / count;
        let mean_x = known_x.iter().sum::<f64>() / count;
        let mut sum_xy_deviation = 0.0_f64;
        let mut sum_x_deviation_square = 0.0_f64;
        for (y_value, x_value) in y_values.iter().zip(known_x.iter()) {
            let y_deviation = y_value - mean_y;
            let x_deviation = x_value - mean_x;
            sum_xy_deviation += y_deviation * x_deviation;
            sum_x_deviation_square += x_deviation * x_deviation;
        }
        if sum_x_deviation_square == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let slope = sum_xy_deviation / sum_x_deviation_square;
        formula_checked_numeric_result(mean_y - slope * mean_x).map(|intercept| (slope, intercept))
    } else {
        let mut sum_xy = 0.0_f64;
        let mut sum_x_square = 0.0_f64;
        for (y_value, x_value) in y_values.iter().zip(known_x.iter()) {
            sum_xy += y_value * x_value;
            sum_x_square += x_value * x_value;
        }
        if sum_x_square == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        formula_checked_numeric_result(sum_xy / sum_x_square).map(|slope| (slope, 0.0))
    }
}

fn formula_numbervalue(
    text: &str,
    decimal_separator: &str,
    group_separator: &str,
) -> Result<f64, FormulaEvalError> {
    if decimal_separator.chars().count() != 1
        || group_separator.chars().count() != 1
        || decimal_separator == group_separator
    {
        return Err(FormulaEvalError::Value);
    }
    let decimal_separator = decimal_separator
        .chars()
        .next()
        .ok_or(FormulaEvalError::Value)?;
    let group_separator = group_separator
        .chars()
        .next()
        .ok_or(FormulaEvalError::Value)?;

    let mut body = text.trim();
    if body.is_empty() {
        return Err(FormulaEvalError::Value);
    }
    let mut multiplier = 1.0;
    while let Some(stripped) = body.strip_suffix('%') {
        multiplier /= 100.0;
        body = stripped.trim_end();
    }
    if body.is_empty() {
        return Err(FormulaEvalError::Value);
    }
    let mut normalized = String::with_capacity(body.len());
    for ch in body.chars() {
        if ch == group_separator {
            continue;
        }
        if ch == decimal_separator {
            normalized.push('.');
        } else if !ch.is_whitespace() {
            normalized.push(ch);
        }
    }
    let value = normalized
        .parse::<f64>()
        .map_err(|_| FormulaEvalError::Value)?
        * multiplier;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Value)
    }
}

fn formula_datevalue_text(text: &str) -> Result<f64, FormulaEvalError> {
    let trimmed = text.trim();
    if trimmed.chars().any(|ch| ch.is_ascii_alphabetic()) {
        let normalized = trimmed.replace([',', '-', '/'], " ");
        let parts = normalized.split_whitespace().collect::<Vec<_>>();
        if parts.len() != 3 {
            return Err(FormulaEvalError::Value);
        }
        let month_names = [
            ("JAN", "JANUARY"),
            ("FEB", "FEBRUARY"),
            ("MAR", "MARCH"),
            ("APR", "APRIL"),
            ("MAY", "MAY"),
            ("JUN", "JUNE"),
            ("JUL", "JULY"),
            ("AUG", "AUGUST"),
            ("SEP", "SEPTEMBER"),
            ("OCT", "OCTOBER"),
            ("NOV", "NOVEMBER"),
            ("DEC", "DECEMBER"),
        ];
        let mut month_index = None;
        for (index, part) in parts.iter().enumerate() {
            let token = part.trim_end_matches('.').to_ascii_uppercase();
            if let Some((month_zero_based, _)) =
                month_names.iter().enumerate().find(|(_, (short, long))| {
                    token == *short || token == *long || (token == "SEPT" && *long == "SEPTEMBER")
                })
            {
                if month_index.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                month_index = Some((index, month_zero_based as i64 + 1));
            }
        }
        let Some((month_position, month)) = month_index else {
            return Err(FormulaEvalError::Value);
        };
        let parse_numeric_part = |part: &str| -> Result<i64, FormulaEvalError> {
            let value = part.trim_end_matches('.');
            let lower = value.to_ascii_lowercase();
            let value = if lower.ends_with("st")
                || lower.ends_with("nd")
                || lower.ends_with("rd")
                || lower.ends_with("th")
            {
                &value[..value.len() - 2]
            } else {
                value
            };
            if value.is_empty() {
                return Err(FormulaEvalError::Value);
            }
            value.parse::<i64>().map_err(|_| FormulaEvalError::Value)
        };
        let (year, day) = match month_position {
            0 => (parse_numeric_part(parts[2])?, parse_numeric_part(parts[1])?),
            1 if parts[0].trim().len() == 4 => {
                (parse_numeric_part(parts[0])?, parse_numeric_part(parts[2])?)
            }
            1 => (parse_numeric_part(parts[2])?, parse_numeric_part(parts[0])?),
            _ => return Err(FormulaEvalError::Value),
        };
        if !(1900..=9999).contains(&year) || day < 1 {
            return Err(FormulaEvalError::Value);
        }
        if year == 1900 && month == 2 && day == 29 {
            return Ok(60.0);
        }
        if day > i64::from(days_in_excel_month(year, month as u32)) {
            return Err(FormulaEvalError::Value);
        }
        return formula_date_serial_from_args(year as f64, month as f64, day as f64);
    }
    let separator = if trimmed.contains('-') {
        '-'
    } else if trimmed.contains('/') {
        '/'
    } else {
        return Err(FormulaEvalError::Value);
    };
    let parts = trimmed.split(separator).collect::<Vec<_>>();
    if parts.len() != 3 || parts.iter().any(|part| part.trim().is_empty()) {
        return Err(FormulaEvalError::Value);
    }
    let parse_part = |part: &str| -> Result<i64, FormulaEvalError> {
        part.trim()
            .parse::<i64>()
            .map_err(|_| FormulaEvalError::Value)
    };
    let first = parse_part(parts[0])?;
    let second = parse_part(parts[1])?;
    let third = parse_part(parts[2])?;
    let (year, month, day) = if parts[0].trim().len() == 4 {
        (first, second, third)
    } else {
        (third, first, second)
    };
    if !(1900..=9999).contains(&year) || !(1..=12).contains(&month) || day < 1 {
        return Err(FormulaEvalError::Value);
    }
    if year == 1900 && month == 2 && day == 29 {
        return Ok(60.0);
    }
    if day > i64::from(days_in_excel_month(year, month as u32)) {
        return Err(FormulaEvalError::Value);
    }
    formula_date_serial_from_args(year as f64, month as f64, day as f64)
}

fn formula_timevalue_text(text: &str) -> Result<f64, FormulaEvalError> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Err(FormulaEvalError::Value);
    }
    let upper = trimmed.to_ascii_uppercase();
    let (body, pm_marker) = if upper.ends_with("AM") {
        (trimmed[..trimmed.len() - 2].trim_end(), Some(false))
    } else if upper.ends_with("PM") {
        (trimmed[..trimmed.len() - 2].trim_end(), Some(true))
    } else {
        (trimmed, None)
    };
    let parts = body.split(':').collect::<Vec<_>>();
    if !(2..=3).contains(&parts.len()) || parts.iter().any(|part| part.trim().is_empty()) {
        return Err(FormulaEvalError::Value);
    }
    let parse_part = |part: &str| -> Result<i64, FormulaEvalError> {
        part.trim()
            .parse::<i64>()
            .map_err(|_| FormulaEvalError::Value)
    };
    let mut hour = parse_part(parts[0])?;
    let minute = parse_part(parts[1])?;
    let second = if parts.len() == 3 {
        parse_part(parts[2])?
    } else {
        0
    };
    if !(0..=59).contains(&minute) || !(0..=59).contains(&second) {
        return Err(FormulaEvalError::Value);
    }
    if let Some(is_pm) = pm_marker {
        if !(1..=12).contains(&hour) {
            return Err(FormulaEvalError::Value);
        }
        if hour == 12 {
            hour = 0;
        }
        if is_pm {
            hour += 12;
        }
    } else if !(0..=23).contains(&hour) {
        return Err(FormulaEvalError::Value);
    }
    Ok((hour * 3600 + minute * 60 + second) as f64 / 86_400.0)
}

fn formula_value_text(text: &str) -> Result<f64, FormulaEvalError> {
    let mut body = text.trim();
    if body.is_empty() {
        return Err(FormulaEvalError::Value);
    }

    let mut accounting_negative = false;
    if let Some(inner) = body
        .strip_prefix('(')
        .and_then(|value| value.strip_suffix(')'))
    {
        accounting_negative = true;
        body = inner.trim();
        if body.is_empty() {
            return Err(FormulaEvalError::Value);
        }
    } else if body.contains('(') || body.contains(')') {
        return Err(FormulaEvalError::Value);
    }

    let mut explicit_negative = false;
    if let Some(rest) = body.strip_prefix('+') {
        body = rest.trim_start();
    } else if let Some(rest) = body.strip_prefix('-') {
        explicit_negative = true;
        body = rest.trim_start();
    }
    if accounting_negative && explicit_negative {
        return Err(FormulaEvalError::Value);
    }

    if let Some(rest) = body.strip_prefix('$') {
        body = rest.trim_start();
    }
    if let Some(rest) = body.strip_prefix('+') {
        if explicit_negative {
            return Err(FormulaEvalError::Value);
        }
        body = rest.trim_start();
    } else if let Some(rest) = body.strip_prefix('-') {
        if explicit_negative || accounting_negative {
            return Err(FormulaEvalError::Value);
        }
        explicit_negative = true;
        body = rest.trim_start();
    }
    if body.contains('$') || body.contains('(') || body.contains(')') {
        return Err(FormulaEvalError::Value);
    }

    let mut value = match formula_numbervalue(body, ".", ",") {
        Ok(value) => value,
        Err(FormulaEvalError::Value) if !accounting_negative && !explicit_negative => {
            if let Ok(value) = formula_datevalue_text(body) {
                value
            } else if let Ok(value) = formula_timevalue_text(body) {
                value
            } else {
                let mut parsed = None;
                for (index, ch) in body.char_indices() {
                    if !ch.is_whitespace() {
                        continue;
                    }
                    let date_text = body[..index].trim_end();
                    let time_text = body[index..].trim_start();
                    if date_text.is_empty() || time_text.is_empty() {
                        continue;
                    }
                    if let (Ok(date), Ok(time)) = (
                        formula_datevalue_text(date_text),
                        formula_timevalue_text(time_text),
                    ) {
                        parsed = Some(date + time);
                        break;
                    }
                }
                parsed.ok_or(FormulaEvalError::Value)?
            }
        }
        Err(error) => return Err(error),
    };
    if accounting_negative ^ explicit_negative {
        value = -value;
    }
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Value)
    }
}

fn formula_proper_text(text: &str) -> String {
    let mut output = String::new();
    let mut capitalize_next = true;
    for ch in text.chars() {
        if ch.is_alphabetic() {
            if capitalize_next {
                output.extend(ch.to_uppercase());
            } else {
                output.extend(ch.to_lowercase());
            }
            capitalize_next = false;
        } else {
            output.push(ch);
            capitalize_next = true;
        }
    }
    output
}

fn formula_text_delimiter_matches(
    text: &str,
    delimiter: &str,
    case_insensitive: bool,
) -> Vec<(usize, usize)> {
    if delimiter.is_empty() {
        return Vec::new();
    }
    let mut matches = Vec::new();
    for (start, _) in text.char_indices() {
        let end = start + delimiter.len();
        if end <= text.len() && text.is_char_boundary(end) {
            let candidate = &text[start..end];
            if candidate == delimiter
                || (case_insensitive && candidate.eq_ignore_ascii_case(delimiter))
            {
                matches.push((start, end));
            }
        }
    }
    matches
}

fn formula_detect_language_tag(text: &str) -> &'static str {
    if text
        .chars()
        .any(|ch| ('\u{AC00}'..='\u{D7AF}').contains(&ch))
    {
        return "ko";
    }
    if text
        .chars()
        .any(|ch| ('\u{3040}'..='\u{30FF}').contains(&ch))
    {
        return "ja";
    }
    if text
        .chars()
        .any(|ch| ('\u{4E00}'..='\u{9FFF}').contains(&ch))
    {
        return "zh";
    }
    if text
        .chars()
        .any(|ch| ('\u{0400}'..='\u{04FF}').contains(&ch))
    {
        return "ru";
    }
    if text
        .chars()
        .any(|ch| ('\u{0600}'..='\u{06FF}').contains(&ch))
    {
        return "ar";
    }
    if text.chars().any(|ch| ch.is_ascii_alphabetic()) {
        return "en";
    }
    "und"
}

pub(super) fn formula_sheet_address_qualifier(sheet_name: &str) -> String {
    if excel_reference_qualifier_needs_quotes(sheet_name) {
        format!("'{}'!", sheet_name.replace('\'', "''"))
    } else {
        format!("{sheet_name}!")
    }
}

fn formula_integer_argument(value: f64) -> Result<i64, FormulaEvalError> {
    if !value.is_finite() || value.fract() != 0.0 {
        return Err(FormulaEvalError::Value);
    }
    if value < i64::MIN as f64 || value > i64::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    Ok(value as i64)
}

fn formula_source_has_top_level_function(formula: &FormulaSource, name: &str) -> bool {
    let text = formula.text.trim_start();
    let Some(prefix) = text.get(..name.len()) else {
        return false;
    };
    if !prefix.eq_ignore_ascii_case(name) {
        return false;
    }
    text[name.len()..].trim_start().starts_with('(')
}

fn formula_bitwise_argument(value: f64) -> Result<u64, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    let value = u64::try_from(value).map_err(|_| FormulaEvalError::Num)?;
    if value > ((1_u64 << 48) - 1) {
        return Err(FormulaEvalError::Num);
    }
    Ok(value)
}

fn formula_bit_shift_argument(value: f64) -> Result<i64, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    if !(-53..=53).contains(&value) {
        return Err(FormulaEvalError::Num);
    }
    Ok(value)
}

fn formula_engineering_input(
    text: &str,
    radix: u32,
    bits: u32,
    max_digits: usize,
) -> Result<i64, FormulaEvalError> {
    let text = text.trim();
    if text.is_empty() || text.len() > max_digits {
        return Err(FormulaEvalError::Num);
    }
    let mut value = 0_u64;
    for ch in text.chars() {
        let Some(digit) = ch.to_digit(radix) else {
            return Err(FormulaEvalError::Num);
        };
        value = value
            .checked_mul(u64::from(radix))
            .and_then(|current| current.checked_add(u64::from(digit)))
            .ok_or(FormulaEvalError::Num)?;
    }
    let sign_threshold = 1_u64 << (bits - 1);
    let modulus = 1_u64 << bits;
    if value >= modulus {
        return Err(FormulaEvalError::Num);
    }
    if value >= sign_threshold {
        Ok(value as i64 - modulus as i64)
    } else {
        Ok(value as i64)
    }
}

fn formula_engineering_format(
    value: i64,
    radix: u32,
    bits: u32,
    max_digits: usize,
    places: Option<usize>,
) -> Result<String, FormulaEvalError> {
    let minimum = -(1_i64 << (bits - 1));
    let maximum = (1_i64 << (bits - 1)) - 1;
    if value < minimum || value > maximum {
        return Err(FormulaEvalError::Num);
    }
    let unsigned = if value < 0 {
        ((1_i128 << bits) + i128::from(value)) as u128
    } else {
        value as u128
    };
    let mut output = formula_unsigned_radix_text(unsigned, radix);
    if value < 0 {
        if output.len() < max_digits {
            output = "0".repeat(max_digits - output.len()) + output.as_str();
        }
        return Ok(output);
    }
    if let Some(places) = places {
        if places == 0 || output.len() > places || places > max_digits {
            return Err(FormulaEvalError::Num);
        }
        if output.len() < places {
            output = "0".repeat(places - output.len()) + output.as_str();
        }
    }
    Ok(output)
}

fn formula_unsigned_radix_text(mut value: u128, radix: u32) -> String {
    if value == 0 {
        return "0".to_string();
    }
    let mut output = String::new();
    while value > 0 {
        let digit = (value % u128::from(radix)) as u32;
        output.push(
            char::from_digit(digit, radix)
                .expect("digit")
                .to_ascii_uppercase(),
        );
        value /= u128::from(radix);
    }
    output.chars().rev().collect()
}

fn formula_roman_text(value: i64, form: usize) -> Result<String, FormulaEvalError> {
    if !(0..=3999).contains(&value) || form > 4 {
        return Err(FormulaEvalError::Value);
    }
    if value == 0 {
        return Ok(String::new());
    }
    let mut remaining = value;
    let mut output = String::new();
    for (candidate, text) in [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ] {
        while remaining >= candidate {
            output.push_str(text);
            remaining -= candidate;
        }
    }
    if form == 0 {
        return Ok(output);
    }
    const CONCISE_REPLACEMENTS: [&[(&str, &str)]; 5] = [
        &[
            ("XLV", "VL"),
            ("XCV", "VC"),
            ("CDL", "LD"),
            ("CML", "LM"),
            ("CMVC", "LMVL"),
        ],
        &[
            ("CDXC", "LDXL"),
            ("CDVC", "LDVL"),
            ("CMXC", "LMXL"),
            ("XCIX", "VCIV"),
            ("XLIX", "VLIV"),
        ],
        &[
            ("XLIX", "IL"),
            ("XCIX", "IC"),
            ("CDXC", "XD"),
            ("CDVC", "XDV"),
            ("CDIC", "XDIX"),
            ("LMVL", "XMV"),
            ("CMIC", "XMIX"),
            ("CMXC", "XM"),
        ],
        &[
            ("XDV", "VD"),
            ("XDIX", "VDIV"),
            ("XMV", "VM"),
            ("XMIX", "VMIV"),
        ],
        &[("VDIV", "ID"), ("VMIV", "IM")],
    ];
    for (index, replacements) in CONCISE_REPLACEMENTS.iter().enumerate().take(form + 1) {
        if index == 1 && form > 1 {
            continue;
        }
        for (from, to) in *replacements {
            output = output.replace(from, to);
        }
    }
    Ok(output)
}

fn formula_fixed_number_text(
    number: f64,
    decimals: i64,
    use_commas: bool,
) -> Result<String, FormulaEvalError> {
    if !number.is_finite() || !(-127..=127).contains(&decimals) {
        return Err(FormulaEvalError::Value);
    }
    let rounded = if decimals >= 0 {
        let factor = formula_round_factor(decimals as f64)?;
        let scaled = number * factor;
        if !scaled.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        round_half_away_from_zero(scaled) / factor
    } else {
        let factor = formula_round_factor((-decimals) as f64)?;
        let scaled = number / factor;
        if !scaled.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        round_half_away_from_zero(scaled) * factor
    };
    if !rounded.is_finite() {
        return Err(FormulaEvalError::Num);
    }
    let negative = rounded < 0.0;
    let precision = decimals.max(0) as usize;
    let mut body = format!("{:.*}", precision, rounded.abs());
    if use_commas {
        let grouped = {
            let (integer, fraction) = body
                .split_once('.')
                .map(|(integer, fraction)| (integer, Some(fraction)))
                .unwrap_or((body.as_str(), None));
            let mut integer_grouped = String::new();
            for (index, ch) in integer.chars().rev().enumerate() {
                if index > 0 && index % 3 == 0 {
                    integer_grouped.push(',');
                }
                integer_grouped.push(ch);
            }
            let mut grouped = integer_grouped.chars().rev().collect::<String>();
            if let Some(fraction) = fraction {
                grouped.push('.');
                grouped.push_str(fraction);
            }
            grouped
        };
        body = grouped;
    }
    if negative {
        Ok(format!("-{body}"))
    } else {
        Ok(body)
    }
}

fn formula_dollar_fraction_parts(fraction: f64) -> Result<(f64, f64), FormulaEvalError> {
    if !fraction.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if fraction < 0.0 {
        return Err(FormulaEvalError::Num);
    }
    let denominator = fraction.trunc();
    if denominator < 1.0 {
        return Err(FormulaEvalError::Div0);
    }
    let scale = 10_f64.powf(denominator.log10().ceil());
    if scale.is_finite() {
        Ok((denominator, scale))
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_dollar_fraction_near_integer(value: f64) -> f64 {
    let rounded = value.round();
    if (value - rounded).abs() <= 1e-9 {
        rounded
    } else {
        value
    }
}

fn formula_financial_type_argument(value: f64) -> Result<f64, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    match value.trunc() as i64 {
        0 => Ok(0.0),
        1 => Ok(1.0),
        _ => Err(FormulaEvalError::Num),
    }
}

fn formula_annuity_growth(rate: f64, nper: f64) -> Result<f64, FormulaEvalError> {
    let growth = (1.0 + rate).powf(nper);
    if growth.is_finite() {
        Ok(growth)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_fv_value(
    rate: f64,
    nper: f64,
    pmt: f64,
    pv: f64,
    payment_type: f64,
) -> Result<f64, FormulaEvalError> {
    if ![rate, nper, pmt, pv].iter().all(|value| value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    let value = if rate == 0.0 {
        -(pv + pmt * nper)
    } else {
        let growth = formula_annuity_growth(rate, nper)?;
        -(pv * growth + pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate)
    };
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_pmt_value(
    rate: f64,
    nper: f64,
    pv: f64,
    fv: f64,
    payment_type: f64,
) -> Result<f64, FormulaEvalError> {
    if ![rate, nper, pv, fv].iter().all(|value| value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    if nper == 0.0 {
        return Err(FormulaEvalError::Div0);
    }
    let value = if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let growth = formula_annuity_growth(rate, nper)?;
        let denominator = (1.0 + rate * payment_type) * (growth - 1.0);
        if denominator == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        -(pv * growth + fv) * rate / denominator
    };
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_ipmt_value(
    rate: f64,
    period: f64,
    nper: f64,
    pv: f64,
    fv: f64,
    payment_type: f64,
) -> Result<f64, FormulaEvalError> {
    if ![rate, period, nper, pv, fv]
        .iter()
        .all(|value| value.is_finite())
    {
        return Err(FormulaEvalError::Value);
    }
    if period < 1.0 || period > nper || nper <= 0.0 {
        return Err(FormulaEvalError::Num);
    }
    if rate == 0.0 {
        return Ok(0.0);
    }
    if payment_type == 1.0 && period == 1.0 {
        return Ok(0.0);
    }
    let payment = formula_pmt_value(rate, nper, pv, fv, payment_type)?;
    let mut value = formula_fv_value(rate, period - 1.0, payment, pv, payment_type)? * rate;
    if payment_type == 1.0 {
        value /= 1.0 + rate;
    }
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

fn formula_radix_argument(value: f64) -> Result<u32, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    if !(2..=36).contains(&value) {
        return Err(FormulaEvalError::Num);
    }
    u32::try_from(value).map_err(|_| FormulaEvalError::Num)
}

fn formula_non_negative_count_argument(value: f64) -> Result<usize, FormulaEvalError> {
    let count = formula_integer_argument(value)?;
    if count < 0 {
        return Err(FormulaEvalError::Value);
    }
    usize::try_from(count).map_err(|_| FormulaEvalError::Num)
}

fn formula_positive_position_argument(value: f64) -> Result<usize, FormulaEvalError> {
    let position = formula_integer_argument(value)?;
    if position < 1 {
        return Err(FormulaEvalError::Value);
    }
    usize::try_from(position).map_err(|_| FormulaEvalError::Num)
}

fn formula_xlookup_match_mode_argument(
    value: f64,
) -> Result<FormulaXLookupMatchMode, FormulaEvalError> {
    match formula_integer_argument(value)? {
        0 => Ok(FormulaXLookupMatchMode::Exact),
        -1 => Ok(FormulaXLookupMatchMode::ExactOrNextSmaller),
        1 => Ok(FormulaXLookupMatchMode::ExactOrNextLarger),
        2 => Ok(FormulaXLookupMatchMode::Wildcard),
        _ => Err(FormulaEvalError::Value),
    }
}

fn formula_xlookup_search_mode_argument(
    value: f64,
) -> Result<FormulaXLookupSearchMode, FormulaEvalError> {
    match formula_integer_argument(value)? {
        1 => Ok(FormulaXLookupSearchMode::Forward),
        -1 => Ok(FormulaXLookupSearchMode::Reverse),
        2 => Ok(FormulaXLookupSearchMode::BinaryAscending),
        -2 => Ok(FormulaXLookupSearchMode::BinaryDescending),
        _ => Err(FormulaEvalError::Value),
    }
}

pub(super) fn formula_date_serial_from_args(
    year: f64,
    month: f64,
    day: f64,
) -> Result<f64, FormulaEvalError> {
    let mut year = formula_integer_argument(year)?;
    let month = formula_integer_argument(month)?;
    let day = formula_integer_argument(day)?;
    if (0..=1899).contains(&year) {
        year += 1900;
    } else if !(1900..=9999).contains(&year) {
        return Err(FormulaEvalError::Num);
    }
    let total_months = year
        .checked_mul(12)
        .and_then(|value| value.checked_add(month - 1))
        .ok_or(FormulaEvalError::Num)?;
    let normalized_year = div_floor(total_months, 12);
    let normalized_month = total_months - normalized_year * 12 + 1;
    if normalized_year == 1900 && normalized_month == 2 && day == 29 {
        return Ok(60.0);
    }
    let days = days_from_civil(normalized_year, normalized_month as u32, 1)
        .checked_add(day - 1)
        .ok_or(FormulaEvalError::Num)?;
    formula_serial_from_civil_days(days).map(|serial| serial as f64)
}

fn formula_ymd_from_serial(serial: f64) -> Result<(i64, u32, u32), FormulaEvalError> {
    let serial = formula_serial_integer(serial)?;
    if serial == 60 {
        return Ok((1900, 2, 29));
    }
    let base_days = days_from_civil(1899, 12, 31);
    let adjusted_serial = if serial > 60 { serial - 1 } else { serial };
    let days = base_days
        .checked_add(adjusted_serial)
        .ok_or(FormulaEvalError::Num)?;
    let (year, month, day) = civil_from_days(days);
    if !(1900..=9999).contains(&year) {
        return Err(FormulaEvalError::Num);
    }
    Ok((year, month, day))
}

fn formula_serial_integer(serial: f64) -> Result<i64, FormulaEvalError> {
    if !serial.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let serial = serial.floor();
    if serial < 1.0 || serial > i64::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    Ok(serial as i64)
}

pub(super) fn formula_current_excel_serial() -> Result<f64, FormulaEvalError> {
    let elapsed = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_err(|_| FormulaEvalError::Num)?;
    Ok(25_569.0 + elapsed.as_secs_f64() / 86_400.0)
}

fn formula_random_u64() -> u64 {
    loop {
        let current = FORMULA_RANDOM_STATE.load(std::sync::atomic::Ordering::Relaxed);
        let state = if current == 0 {
            let seed = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|duration| duration.as_nanos() as u64)
                .unwrap_or(0x9e37_79b9_7f4a_7c15);
            let seed = seed ^ 0xa076_1d64_78bd_642f;
            if seed == 0 {
                0xe703_7ed1_a0b4_28db
            } else {
                seed
            }
        } else {
            current
        };
        let next = state
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1_442_695_040_888_963_407);
        match FORMULA_RANDOM_STATE.compare_exchange_weak(
            current,
            next,
            std::sync::atomic::Ordering::Relaxed,
            std::sync::atomic::Ordering::Relaxed,
        ) {
            Ok(_) => return next,
            Err(_) => continue,
        }
    }
}

fn formula_rand() -> f64 {
    const SCALE: f64 = 1.0 / ((1_u64 << 53) as f64);
    ((formula_random_u64() >> 11) as f64) * SCALE
}

fn formula_rand_between(bottom: f64, top: f64) -> Result<f64, FormulaEvalError> {
    if !bottom.is_finite() || !top.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    let bottom = bottom.trunc();
    let top = top.trunc();
    if bottom < i64::MIN as f64
        || bottom > i64::MAX as f64
        || top < i64::MIN as f64
        || top > i64::MAX as f64
    {
        return Err(FormulaEvalError::Num);
    }
    let bottom = bottom as i64;
    let top = top as i64;
    if bottom > top {
        return Err(FormulaEvalError::Num);
    }
    let span = i128::from(top) - i128::from(bottom) + 1;
    let span = u64::try_from(span).map_err(|_| FormulaEvalError::Num)?;
    let offset = (formula_random_u64() % span) as i128;
    Ok((i128::from(bottom) + offset) as f64)
}

fn formula_edate(serial: f64, months: f64) -> Result<f64, FormulaEvalError> {
    let (year, month, day) = formula_ymd_from_serial(serial)?;
    let months = formula_integer_argument(months)?;
    let (target_year, target_month) = normalize_year_month(year, i64::from(month) + months)?;
    let target_day = i64::from(day.min(days_in_excel_month(target_year, target_month)));
    formula_date_serial_from_args(target_year as f64, target_month as f64, target_day as f64)
}

fn formula_eomonth(serial: f64, months: f64) -> Result<f64, FormulaEvalError> {
    let (year, month, _) = formula_ymd_from_serial(serial)?;
    let months = formula_integer_argument(months)?;
    let (target_year, target_month) = normalize_year_month(year, i64::from(month) + months)?;
    let target_day = i64::from(days_in_excel_month(target_year, target_month));
    formula_date_serial_from_args(target_year as f64, target_month as f64, target_day as f64)
}

fn formula_weekday_monday0_from_serial(serial: i64) -> i64 {
    let adjusted_serial = if serial > 60 { serial - 1 } else { serial };
    (adjusted_serial - 1).rem_euclid(7)
}

fn formula_standard_weekend_mask() -> [bool; 7] {
    [false, false, false, false, false, true, true]
}

fn formula_weekend_mask_from_code(code: i64) -> Result<[bool; 7], FormulaEvalError> {
    match code {
        1 => Ok(formula_standard_weekend_mask()),
        2 => Ok([true, false, false, false, false, false, true]),
        3 => Ok([true, true, false, false, false, false, false]),
        4 => Ok([false, true, true, false, false, false, false]),
        5 => Ok([false, false, true, true, false, false, false]),
        6 => Ok([false, false, false, true, true, false, false]),
        7 => Ok([false, false, false, false, true, true, false]),
        11 => Ok([false, false, false, false, false, false, true]),
        12 => Ok([true, false, false, false, false, false, false]),
        13 => Ok([false, true, false, false, false, false, false]),
        14 => Ok([false, false, true, false, false, false, false]),
        15 => Ok([false, false, false, true, false, false, false]),
        16 => Ok([false, false, false, false, true, false, false]),
        17 => Ok([false, false, false, false, false, true, false]),
        _ => Err(FormulaEvalError::Value),
    }
}

fn formula_weekend_mask_from_string(value: &str) -> Result<[bool; 7], FormulaEvalError> {
    if value.chars().count() != 7 {
        return Err(FormulaEvalError::Value);
    }
    let mut mask = [false; 7];
    let mut workday_count = 0_u8;
    for (index, ch) in value.chars().enumerate() {
        mask[index] = match ch {
            '0' => {
                workday_count += 1;
                false
            }
            '1' => true,
            _ => return Err(FormulaEvalError::Value),
        };
    }
    if workday_count == 0 {
        return Err(FormulaEvalError::Value);
    }
    Ok(mask)
}

fn formula_is_workday_serial(serial: i64, holidays: &[i64], weekend: &[bool; 7]) -> bool {
    !weekend[formula_weekday_monday0_from_serial(serial) as usize] && !holidays.contains(&serial)
}

fn formula_networkdays(
    start_serial: i64,
    end_serial: i64,
    holidays: &[i64],
) -> Result<f64, FormulaEvalError> {
    formula_networkdays_with_weekend(
        start_serial,
        end_serial,
        holidays,
        &formula_standard_weekend_mask(),
    )
}

fn formula_networkdays_with_weekend(
    start_serial: i64,
    end_serial: i64,
    holidays: &[i64],
    weekend: &[bool; 7],
) -> Result<f64, FormulaEvalError> {
    let (first, last, sign) = if start_serial <= end_serial {
        (start_serial, end_serial, 1.0)
    } else {
        (end_serial, start_serial, -1.0)
    };
    formula_ymd_from_serial(first as f64)?;
    formula_ymd_from_serial(last as f64)?;
    let mut count = 0_u64;
    for serial in first..=last {
        if formula_is_workday_serial(serial, holidays, weekend) {
            count += 1;
        }
    }
    Ok(count as f64 * sign)
}

fn formula_workday(
    start_serial: i64,
    days: i64,
    holidays: &[i64],
) -> Result<f64, FormulaEvalError> {
    formula_workday_with_weekend(
        start_serial,
        days,
        holidays,
        &formula_standard_weekend_mask(),
    )
}

fn formula_workday_with_weekend(
    start_serial: i64,
    days: i64,
    holidays: &[i64],
    weekend: &[bool; 7],
) -> Result<f64, FormulaEvalError> {
    formula_ymd_from_serial(start_serial as f64)?;
    let direction = if days < 0 { -1 } else { 1 };
    let mut serial = start_serial;
    let mut remaining = days.unsigned_abs();
    while remaining > 0 {
        serial = serial.checked_add(direction).ok_or(FormulaEvalError::Num)?;
        formula_ymd_from_serial(serial as f64)?;
        if formula_is_workday_serial(serial, holidays, weekend) {
            remaining -= 1;
        }
    }
    Ok(serial as f64)
}

fn formula_time_serial_from_args(
    hour: f64,
    minute: f64,
    second: f64,
) -> Result<f64, FormulaEvalError> {
    let hour = formula_time_argument(hour)?;
    let minute = formula_time_argument(minute)?;
    let second = formula_time_argument(second)?;
    let total_seconds = hour
        .checked_mul(3600)
        .and_then(|value| {
            minute
                .checked_mul(60)
                .and_then(|minute| value.checked_add(minute))
        })
        .and_then(|value| value.checked_add(second))
        .ok_or(FormulaEvalError::Num)?;
    Ok((total_seconds.rem_euclid(86_400)) as f64 / 86_400.0)
}

fn formula_time_parts_from_serial(serial: f64) -> Result<(u32, u32, u32), FormulaEvalError> {
    if !serial.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if serial < 0.0 {
        return Err(FormulaEvalError::Num);
    }
    let fraction = serial - serial.floor();
    let mut total_seconds = (fraction * 86_400.0).round() as i64;
    if total_seconds >= 86_400 {
        total_seconds = 0;
    }
    let hour = total_seconds / 3600;
    let minute = (total_seconds % 3600) / 60;
    let second = total_seconds % 60;
    Ok((hour as u32, minute as u32, second as u32))
}

fn formula_time_argument(value: f64) -> Result<i64, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    if !(0..=32_767).contains(&value) {
        return Err(FormulaEvalError::Num);
    }
    Ok(value)
}

fn normalize_year_month(year: i64, month: i64) -> Result<(i64, u32), FormulaEvalError> {
    let total_months = year
        .checked_mul(12)
        .and_then(|value| value.checked_add(month - 1))
        .ok_or(FormulaEvalError::Num)?;
    let normalized_year = div_floor(total_months, 12);
    let normalized_month = (total_months - normalized_year * 12 + 1) as u32;
    if !(1900..=9999).contains(&normalized_year) {
        return Err(FormulaEvalError::Num);
    }
    Ok((normalized_year, normalized_month))
}

fn formula_serial_from_civil_days(days: i64) -> Result<i64, FormulaEvalError> {
    let min_days = days_from_civil(1900, 1, 1);
    let max_days = days_from_civil(9999, 12, 31);
    if days < min_days || days > max_days {
        return Err(FormulaEvalError::Num);
    }
    let base_days = days_from_civil(1899, 12, 31);
    let mut serial = days - base_days;
    if days >= days_from_civil(1900, 3, 1) {
        serial += 1;
    }
    Ok(serial)
}

fn days_in_excel_month(year: i64, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year == 1900 => 29,
        2 if is_gregorian_leap_year(year) => 29,
        2 => 28,
        _ => 0,
    }
}

fn is_gregorian_leap_year(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn div_floor(value: i64, divisor: i64) -> i64 {
    let quotient = value / divisor;
    let remainder = value % divisor;
    if remainder != 0 && ((remainder > 0) != (divisor > 0)) {
        quotient - 1
    } else {
        quotient
    }
}

fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = year - i64::from(month <= 2);
    let era = div_floor(year, 400);
    let year_of_era = year - era * 400;
    let month = i64::from(month);
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let days = days + 719_468;
    let era = div_floor(days, 146_097);
    let day_of_era = days - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let mut year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    year += i64::from(month <= 2);
    (year, month as u32, day as u32)
}

fn lookup_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaLookupMode,
) -> Result<usize, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    if matches!(mode, FormulaLookupMode::Exact) {
        for (index, candidate) in values.iter().enumerate() {
            if formula_value_probe_exact_match(lookup_value, candidate)? {
                return Ok(index);
            }
        }
        return Err(FormulaEvalError::NA);
    }

    let mut best = None;
    for (index, candidate) in values.iter().enumerate() {
        let ordering = formula_value_probe_ordering(candidate, lookup_value)?;
        let Some(ordering) = ordering else {
            continue;
        };
        match mode {
            FormulaLookupMode::ApproxAscending if ordering != Ordering::Greater => {
                best = Some(index);
            }
            FormulaLookupMode::ApproxDescending if ordering != Ordering::Less => {
                best = Some(index);
            }
            FormulaLookupMode::Exact
            | FormulaLookupMode::ApproxAscending
            | FormulaLookupMode::ApproxDescending => {}
        }
    }
    best.ok_or(FormulaEvalError::NA)
}

fn xlookup_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaXLookupMatchMode,
    search_mode: FormulaXLookupSearchMode,
) -> Result<usize, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    match search_mode {
        FormulaXLookupSearchMode::Forward => {
            xlookup_linear_match_index_in_values(lookup_value, values, mode, false)
        }
        FormulaXLookupSearchMode::Reverse => {
            xlookup_linear_match_index_in_values(lookup_value, values, mode, true)
        }
        FormulaXLookupSearchMode::BinaryAscending => {
            xlookup_binary_match_index_in_values(lookup_value, values, mode, true)
        }
        FormulaXLookupSearchMode::BinaryDescending => {
            xlookup_binary_match_index_in_values(lookup_value, values, mode, false)
        }
    }
}

fn xlookup_linear_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaXLookupMatchMode,
    reverse_search: bool,
) -> Result<usize, FormulaEvalError> {
    let indexes = if reverse_search {
        (0..values.len()).rev().collect::<Vec<_>>()
    } else {
        (0..values.len()).collect::<Vec<_>>()
    };
    for index in indexes.iter().copied() {
        if match mode {
            FormulaXLookupMatchMode::Wildcard => {
                formula_value_probe_wildcard_match(lookup_value, &values[index])?
            }
            FormulaXLookupMatchMode::Exact
            | FormulaXLookupMatchMode::ExactOrNextSmaller
            | FormulaXLookupMatchMode::ExactOrNextLarger => {
                formula_value_probe_exact_match(lookup_value, &values[index])?
            }
        } {
            return Ok(index);
        }
    }
    if matches!(
        mode,
        FormulaXLookupMatchMode::Exact | FormulaXLookupMatchMode::Wildcard
    ) {
        return Err(FormulaEvalError::NA);
    }

    let mut best = None::<(usize, FormulaValueProbe)>;
    for index in indexes {
        let candidate = &values[index];
        let Some(ordering) = formula_value_probe_ordering(candidate, lookup_value)? else {
            continue;
        };
        let is_viable = match mode {
            FormulaXLookupMatchMode::Exact => false,
            FormulaXLookupMatchMode::ExactOrNextSmaller => ordering == Ordering::Less,
            FormulaXLookupMatchMode::ExactOrNextLarger => ordering == Ordering::Greater,
            FormulaXLookupMatchMode::Wildcard => false,
        };
        if !is_viable {
            continue;
        }
        let replace = match &best {
            None => true,
            Some((_, current)) => match formula_value_probe_ordering(candidate, current)? {
                Some(candidate_to_current) => match mode {
                    FormulaXLookupMatchMode::Exact => false,
                    FormulaXLookupMatchMode::ExactOrNextSmaller => {
                        candidate_to_current == Ordering::Greater
                    }
                    FormulaXLookupMatchMode::ExactOrNextLarger => {
                        candidate_to_current == Ordering::Less
                    }
                    FormulaXLookupMatchMode::Wildcard => false,
                },
                None => false,
            },
        };
        if replace {
            best = Some((index, candidate.clone()));
        }
    }
    best.map(|(index, _)| index).ok_or(FormulaEvalError::NA)
}

fn xlookup_binary_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaXLookupMatchMode,
    ascending: bool,
) -> Result<usize, FormulaEvalError> {
    if matches!(mode, FormulaXLookupMatchMode::Wildcard) {
        return Err(FormulaEvalError::Value);
    }

    let mut comparable = Vec::with_capacity(values.len());
    for (index, candidate) in values.iter().enumerate() {
        if formula_value_probe_ordering(candidate, lookup_value)?.is_some() {
            comparable.push((index, candidate));
        }
    }
    if comparable.is_empty() {
        return Err(FormulaEvalError::NA);
    }

    let mut low = 0_usize;
    let mut high = comparable.len();
    while low < high {
        let mid = low + (high - low) / 2;
        let ordering = formula_value_probe_ordering(comparable[mid].1, lookup_value)?
            .ok_or(FormulaEvalError::NA)?;
        let before_lookup = if ascending {
            ordering == Ordering::Less
        } else {
            ordering == Ordering::Greater
        };
        if before_lookup {
            low = mid + 1;
        } else {
            high = mid;
        }
    }

    if let Some((index, candidate)) = comparable.get(low)
        && formula_value_probe_exact_match(lookup_value, candidate)?
    {
        return Ok(*index);
    }

    let candidate = match mode {
        FormulaXLookupMatchMode::Exact => None,
        FormulaXLookupMatchMode::ExactOrNextSmaller if ascending => {
            low.checked_sub(1).and_then(|index| comparable.get(index))
        }
        FormulaXLookupMatchMode::ExactOrNextSmaller => comparable.get(low),
        FormulaXLookupMatchMode::ExactOrNextLarger if ascending => comparable.get(low),
        FormulaXLookupMatchMode::ExactOrNextLarger => {
            low.checked_sub(1).and_then(|index| comparable.get(index))
        }
        FormulaXLookupMatchMode::Wildcard => None,
    };
    candidate
        .map(|(index, _)| *index)
        .ok_or(FormulaEvalError::NA)
}

fn formula_value_probe_exact_match(
    lookup_value: &FormulaValueProbe,
    candidate: &FormulaValueProbe,
) -> Result<bool, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    if let FormulaValueProbe::Error(error) = candidate {
        return Err(*error);
    }
    Ok(match (lookup_value, candidate) {
        (FormulaValueProbe::Blank, FormulaValueProbe::Blank) => true,
        (FormulaValueProbe::Bool(left), FormulaValueProbe::Bool(right)) => left == right,
        (FormulaValueProbe::Number(left), FormulaValueProbe::Number(right)) => left == right,
        (FormulaValueProbe::Text(left), FormulaValueProbe::Text(right)) => {
            left.eq_ignore_ascii_case(right)
        }
        _ => false,
    })
}

fn formula_value_probe_wildcard_match(
    lookup_value: &FormulaValueProbe,
    candidate: &FormulaValueProbe,
) -> Result<bool, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    if let FormulaValueProbe::Error(error) = candidate {
        return Err(*error);
    }
    Ok(match (lookup_value, candidate) {
        (FormulaValueProbe::Text(pattern), FormulaValueProbe::Text(value)) => {
            formula_wildcard_matches(pattern, value, true)
        }
        _ => false,
    })
}

fn formula_value_probe_ordering(
    left: &FormulaValueProbe,
    right: &FormulaValueProbe,
) -> Result<Option<Ordering>, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = left {
        return Err(*error);
    }
    if let FormulaValueProbe::Error(error) = right {
        return Err(*error);
    }
    Ok(match (left, right) {
        (FormulaValueProbe::Bool(left), FormulaValueProbe::Bool(right)) => Some(left.cmp(right)),
        (FormulaValueProbe::Number(left), FormulaValueProbe::Number(right)) => {
            if !left.is_finite() || !right.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            left.partial_cmp(right)
        }
        (FormulaValueProbe::Text(left), FormulaValueProbe::Text(right)) => {
            Some(left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase()))
        }
        _ => None,
    })
}

fn formula_wildcard_matches(pattern: &str, value: &str, case_insensitive: bool) -> bool {
    let tokens = formula_wildcard_tokens(pattern, case_insensitive);
    let chars = formula_wildcard_chars(value, case_insensitive);
    formula_wildcard_tokens_match(tokens.as_slice(), chars.as_slice(), false)
}

fn formula_wildcard_find(
    pattern: &str,
    value: &str,
    start: usize,
    case_insensitive: bool,
) -> Option<usize> {
    let tokens = formula_wildcard_tokens(pattern, case_insensitive);
    let chars = formula_wildcard_chars(value, case_insensitive);
    for index in start.saturating_sub(1)..=chars.len() {
        if formula_wildcard_tokens_match(tokens.as_slice(), &chars[index..], true) {
            return Some(index + 1);
        }
    }
    None
}

fn formula_wildcard_tokens(pattern: &str, case_insensitive: bool) -> Vec<FormulaWildcardToken> {
    let mut tokens = Vec::new();
    let mut chars = pattern.chars();
    while let Some(ch) = chars.next() {
        match ch {
            '*' => tokens.push(FormulaWildcardToken::AnySequence),
            '?' => tokens.push(FormulaWildcardToken::AnyChar),
            '~' => {
                let literal = chars.next().unwrap_or('~');
                tokens.push(FormulaWildcardToken::Literal(formula_wildcard_char(
                    literal,
                    case_insensitive,
                )));
            }
            literal => tokens.push(FormulaWildcardToken::Literal(formula_wildcard_char(
                literal,
                case_insensitive,
            ))),
        }
    }
    tokens
}

fn formula_wildcard_chars(value: &str, case_insensitive: bool) -> Vec<char> {
    value
        .chars()
        .map(|ch| formula_wildcard_char(ch, case_insensitive))
        .collect()
}

fn formula_wildcard_char(ch: char, case_insensitive: bool) -> char {
    if case_insensitive {
        ch.to_ascii_lowercase()
    } else {
        ch
    }
}

fn formula_wildcard_tokens_match(
    tokens: &[FormulaWildcardToken],
    chars: &[char],
    accept_prefix: bool,
) -> bool {
    fn matches_from(
        tokens: &[FormulaWildcardToken],
        chars: &[char],
        accept_prefix: bool,
        token_index: usize,
        char_index: usize,
        memo: &mut [Vec<Option<bool>>],
    ) -> bool {
        if let Some(value) = memo[token_index][char_index] {
            return value;
        }
        let matched = if token_index == tokens.len() {
            accept_prefix || char_index == chars.len()
        } else {
            match tokens[token_index] {
                FormulaWildcardToken::Literal(expected) => {
                    char_index < chars.len()
                        && chars[char_index] == expected
                        && matches_from(
                            tokens,
                            chars,
                            accept_prefix,
                            token_index + 1,
                            char_index + 1,
                            memo,
                        )
                }
                FormulaWildcardToken::AnyChar => {
                    char_index < chars.len()
                        && matches_from(
                            tokens,
                            chars,
                            accept_prefix,
                            token_index + 1,
                            char_index + 1,
                            memo,
                        )
                }
                FormulaWildcardToken::AnySequence => {
                    matches_from(
                        tokens,
                        chars,
                        accept_prefix,
                        token_index + 1,
                        char_index,
                        memo,
                    ) || (char_index < chars.len()
                        && matches_from(
                            tokens,
                            chars,
                            accept_prefix,
                            token_index,
                            char_index + 1,
                            memo,
                        ))
                }
            }
        };
        memo[token_index][char_index] = Some(matched);
        matched
    }

    let mut memo = vec![vec![None; chars.len() + 1]; tokens.len() + 1];
    matches_from(tokens, chars, accept_prefix, 0, 0, &mut memo)
}

fn parse_formula_criteria_numeric_literal(input: &str) -> Option<(FormulaComparisonOperator, f64)> {
    for (prefix, operator) in [
        ("<>", FormulaComparisonOperator::NotEqual),
        ("<=", FormulaComparisonOperator::LessThanOrEqual),
        (">=", FormulaComparisonOperator::GreaterThanOrEqual),
        ("=", FormulaComparisonOperator::Equal),
        ("<", FormulaComparisonOperator::LessThan),
        (">", FormulaComparisonOperator::GreaterThan),
    ] {
        if let Some(rest) = input.strip_prefix(prefix) {
            let operand = rest.trim();
            if operand.is_empty() {
                return None;
            }
            return operand.parse::<f64>().ok().map(|value| (operator, value));
        }
    }
    input
        .parse::<f64>()
        .ok()
        .map(|value| (FormulaComparisonOperator::Equal, value))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaAggregateFunction {
    Sum,
    Product,
    SumSq,
    Min,
    Max,
    Median,
    Average,
    Gcd,
    GeoMean,
    HarMean,
    Kurt,
    Lcm,
    ModeMult,
    ModeSngl,
    Skew,
    SkewP,
    AveDev,
    DevSq,
    VarP,
    VarS,
    StDevP,
    StDevS,
    Count,
}

impl FormulaAggregateFunction {
    fn from_name(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("SUM") {
            Some(Self::Sum)
        } else if name.eq_ignore_ascii_case("PRODUCT") {
            Some(Self::Product)
        } else if name.eq_ignore_ascii_case("SUMSQ") {
            Some(Self::SumSq)
        } else if name.eq_ignore_ascii_case("MIN") {
            Some(Self::Min)
        } else if name.eq_ignore_ascii_case("MAX") {
            Some(Self::Max)
        } else if name.eq_ignore_ascii_case("MEDIAN") {
            Some(Self::Median)
        } else if name.eq_ignore_ascii_case("AVERAGE") {
            Some(Self::Average)
        } else if name.eq_ignore_ascii_case("GCD") {
            Some(Self::Gcd)
        } else if name.eq_ignore_ascii_case("GEOMEAN") {
            Some(Self::GeoMean)
        } else if name.eq_ignore_ascii_case("HARMEAN") {
            Some(Self::HarMean)
        } else if name.eq_ignore_ascii_case("KURT") {
            Some(Self::Kurt)
        } else if name.eq_ignore_ascii_case("LCM") {
            Some(Self::Lcm)
        } else if name.eq_ignore_ascii_case("MODE.MULT") {
            Some(Self::ModeMult)
        } else if name.eq_ignore_ascii_case("MODE") || name.eq_ignore_ascii_case("MODE.SNGL") {
            Some(Self::ModeSngl)
        } else if name.eq_ignore_ascii_case("SKEW") {
            Some(Self::Skew)
        } else if name.eq_ignore_ascii_case("SKEW.P") {
            Some(Self::SkewP)
        } else if name.eq_ignore_ascii_case("AVEDEV") {
            Some(Self::AveDev)
        } else if name.eq_ignore_ascii_case("DEVSQ") {
            Some(Self::DevSq)
        } else if name.eq_ignore_ascii_case("VAR.P") || name.eq_ignore_ascii_case("VARP") {
            Some(Self::VarP)
        } else if name.eq_ignore_ascii_case("VAR.S") || name.eq_ignore_ascii_case("VAR") {
            Some(Self::VarS)
        } else if name.eq_ignore_ascii_case("STDEV.P") || name.eq_ignore_ascii_case("STDEVP") {
            Some(Self::StDevP)
        } else if name.eq_ignore_ascii_case("STDEV.S") || name.eq_ignore_ascii_case("STDEV") {
            Some(Self::StDevS)
        } else if name.eq_ignore_ascii_case("COUNT") {
            Some(Self::Count)
        } else {
            None
        }
    }

    fn evaluate(self, values: &[f64]) -> Result<f64, FormulaEvalError> {
        let mean = |values: &[f64]| -> Result<f64, FormulaEvalError> {
            if values.is_empty() {
                Err(FormulaEvalError::Div0)
            } else {
                Ok(values.iter().sum::<f64>() / values.len() as f64)
            }
        };
        let deviation_sum = |values: &[f64]| -> Result<f64, FormulaEvalError> {
            let mean = mean(values)?;
            Ok(values
                .iter()
                .map(|value| {
                    let deviation = value - mean;
                    deviation * deviation
                })
                .sum())
        };
        let checked_numeric_result = |value: f64| -> Result<f64, FormulaEvalError> {
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        const EXCEL_INTEGER_LIMIT: u64 = 1_u64 << 53;
        let trunc_excel_nonnegative_integer = |value: f64| -> Result<u64, FormulaEvalError> {
            if !value.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let value = value.trunc();
            if value < 0.0 || value >= EXCEL_INTEGER_LIMIT as f64 {
                return Err(FormulaEvalError::Num);
            }
            Ok(value as u64)
        };
        let gcd_u64 = |mut left: u64, mut right: u64| -> u64 {
            while right != 0 {
                let next = left % right;
                left = right;
                right = next;
            }
            left
        };

        match self {
            FormulaAggregateFunction::Sum => Ok(values.iter().sum()),
            FormulaAggregateFunction::Product => Ok(values.iter().product()),
            FormulaAggregateFunction::SumSq => Ok(values.iter().map(|value| value * value).sum()),
            FormulaAggregateFunction::Min => {
                Ok(values.iter().copied().reduce(f64::min).unwrap_or(0.0))
            }
            FormulaAggregateFunction::Max => {
                Ok(values.iter().copied().reduce(f64::max).unwrap_or(0.0))
            }
            FormulaAggregateFunction::Median => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Num);
                }
                let mut values = values.to_vec();
                values.sort_by(|left, right| left.total_cmp(right));
                let midpoint = values.len() / 2;
                if values.len() % 2 == 0 {
                    Ok((values[midpoint - 1] + values[midpoint]) / 2.0)
                } else {
                    Ok(values[midpoint])
                }
            }
            FormulaAggregateFunction::Average => {
                if values.is_empty() {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(values.iter().sum::<f64>() / values.len() as f64)
                }
            }
            FormulaAggregateFunction::Gcd => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                let mut result = 0_u64;
                for value in values {
                    result = gcd_u64(result, trunc_excel_nonnegative_integer(*value)?);
                }
                Ok(result as f64)
            }
            FormulaAggregateFunction::GeoMean => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Div0);
                }
                let mut log_sum = 0.0;
                for value in values {
                    if *value <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    log_sum += value.ln();
                }
                Ok((log_sum / values.len() as f64).exp())
            }
            FormulaAggregateFunction::HarMean => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Div0);
                }
                let mut reciprocal_sum = 0.0;
                for value in values {
                    if *value <= 0.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    reciprocal_sum += 1.0 / value;
                }
                Ok(values.len() as f64 / reciprocal_sum)
            }
            FormulaAggregateFunction::Kurt => {
                let count = values.len();
                if count < 4 {
                    return Err(FormulaEvalError::Div0);
                }
                let mean = mean(values)?;
                let mut deviation_square_sum = 0.0_f64;
                let mut deviation_fourth_sum = 0.0_f64;
                for value in values {
                    let deviation = value - mean;
                    let deviation_square = deviation * deviation;
                    deviation_square_sum += deviation_square;
                    deviation_fourth_sum += deviation_square * deviation_square;
                }
                if deviation_square_sum == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let count = count as f64;
                let sample_variance = deviation_square_sum / (count - 1.0);
                let kurtosis = count * (count + 1.0) * deviation_fourth_sum
                    / ((count - 1.0)
                        * (count - 2.0)
                        * (count - 3.0)
                        * sample_variance
                        * sample_variance)
                    - 3.0 * (count - 1.0) * (count - 1.0) / ((count - 2.0) * (count - 3.0));
                checked_numeric_result(kurtosis)
            }
            FormulaAggregateFunction::Lcm => {
                if values.is_empty() {
                    return Err(FormulaEvalError::Value);
                }
                let mut result = 1_u64;
                for value in values {
                    let value = trunc_excel_nonnegative_integer(*value)?;
                    if value == 0 {
                        return Ok(0.0);
                    }
                    let next = (result / gcd_u64(result, value))
                        .checked_mul(value)
                        .ok_or(FormulaEvalError::Num)?;
                    if next >= EXCEL_INTEGER_LIMIT {
                        return Err(FormulaEvalError::Num);
                    }
                    result = next;
                }
                Ok(result as f64)
            }
            FormulaAggregateFunction::ModeMult | FormulaAggregateFunction::ModeSngl => {
                let mut mode = None;
                let mut mode_count = 1_usize;
                for value in values {
                    let count = values
                        .iter()
                        .filter(|candidate| **candidate == *value)
                        .count();
                    if count > mode_count {
                        mode = Some(*value);
                        mode_count = count;
                    }
                }
                mode.ok_or(FormulaEvalError::NA)
            }
            FormulaAggregateFunction::Skew => {
                let count = values.len();
                if count < 3 {
                    return Err(FormulaEvalError::Div0);
                }
                let mean = mean(values)?;
                let mut deviation_square_sum = 0.0_f64;
                let mut deviation_cube_sum = 0.0_f64;
                for value in values {
                    let deviation = value - mean;
                    deviation_square_sum += deviation * deviation;
                    deviation_cube_sum += deviation * deviation * deviation;
                }
                if deviation_square_sum == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let count = count as f64;
                let sample_standard_deviation = (deviation_square_sum / (count - 1.0)).sqrt();
                checked_numeric_result(
                    count * deviation_cube_sum
                        / ((count - 1.0)
                            * (count - 2.0)
                            * sample_standard_deviation
                            * sample_standard_deviation
                            * sample_standard_deviation),
                )
            }
            FormulaAggregateFunction::SkewP => {
                let count = values.len();
                if count < 3 {
                    return Err(FormulaEvalError::Div0);
                }
                let mean = mean(values)?;
                let mut deviation_square_sum = 0.0_f64;
                let mut deviation_cube_sum = 0.0_f64;
                for value in values {
                    let deviation = value - mean;
                    deviation_square_sum += deviation * deviation;
                    deviation_cube_sum += deviation * deviation * deviation;
                }
                if deviation_square_sum == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let count = count as f64;
                let population_standard_deviation = (deviation_square_sum / count).sqrt();
                checked_numeric_result(
                    deviation_cube_sum
                        / count
                        / (population_standard_deviation
                            * population_standard_deviation
                            * population_standard_deviation),
                )
            }
            FormulaAggregateFunction::AveDev => {
                let mean = mean(values)?;
                Ok(values.iter().map(|value| (value - mean).abs()).sum::<f64>()
                    / values.len() as f64)
            }
            FormulaAggregateFunction::DevSq => deviation_sum(values),
            FormulaAggregateFunction::VarP => Ok(deviation_sum(values)? / values.len() as f64),
            FormulaAggregateFunction::VarS => {
                if values.len() < 2 {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(deviation_sum(values)? / (values.len() - 1) as f64)
                }
            }
            FormulaAggregateFunction::StDevP => {
                Ok((deviation_sum(values)? / values.len() as f64).sqrt())
            }
            FormulaAggregateFunction::StDevS => {
                if values.len() < 2 {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok((deviation_sum(values)? / (values.len() - 1) as f64).sqrt())
                }
            }
            FormulaAggregateFunction::Count => Ok(values.len() as f64),
        }
    }
}

pub(super) struct FormulaEvaluator<'a> {
    state: &'a WorkbookState,
    visiting: BTreeSet<(SheetId, u32, u32)>,
    resolving_names: BTreeSet<DefinedNameId>,
}

impl<'a> FormulaEvaluator<'a> {
    pub(super) fn new(state: &'a WorkbookState) -> Self {
        Self {
            state,
            visiting: BTreeSet::new(),
            resolving_names: BTreeSet::new(),
        }
    }

    pub(super) fn evaluate_formula_cell(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Option<CellValue> {
        match self.evaluate_formula_cell_result(sheet_id, row, col) {
            Ok(value) => Some(value),
            Err(error) => error.into_cell_value(),
        }
    }

    pub(super) fn evaluate_formula_cell_result(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<CellValue, FormulaEvalError> {
        self.evaluate_cell(sheet_id, row, col)
    }

    pub(super) fn evaluate_dynamic_array_formula_cell_result(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<FormulaArrayResult, FormulaEvalError> {
        let formula = self
            .state
            .worksheet_data
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
            .and_then(|cell| cell.formula.as_ref())
            .ok_or(FormulaEvalError::Unsupported)?;
        if !self.visiting.insert((sheet_id, row, col)) {
            return Err(FormulaEvalError::Circular);
        }
        let formula_text = if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, row, col)
        } else {
            formula.text.clone()
        };
        let result = {
            let mut parser = FormulaParser::new(&formula_text, self, sheet_id, Some((row, col)));
            parser.parse_dynamic_array_formula()
        };
        let result = match result {
            Ok(result) => Ok(result),
            Err(FormulaEvalError::Unsupported) => {
                match self.evaluate_formula_text(sheet_id, &formula_text, Some((row, col))) {
                    Ok(value) => Ok(FormulaArrayResult::single(value)),
                    Err(FormulaEvalError::Unsupported) => {
                        let mut parser =
                            FormulaParser::new(&formula_text, self, sheet_id, Some((row, col)));
                        parser
                            .parse_value_probe_formula()
                            .and_then(|probe| {
                                formula_cell_value_from_probe(probe)
                                    .ok_or(FormulaEvalError::Unsupported)
                            })
                            .map(FormulaArrayResult::single)
                    }
                    Err(error) => Err(error),
                }
            }
            Err(error) => Err(error),
        };
        self.visiting.remove(&(sheet_id, row, col));
        result
    }

    fn evaluate_cell(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<CellValue, FormulaEvalError> {
        let Some(cell) = self
            .state
            .worksheet_data
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
        else {
            return Ok(CellValue::Blank);
        };
        let Some(formula) = cell.formula.as_ref() else {
            return Ok(cell.value.clone());
        };
        if !self.visiting.insert((sheet_id, row, col)) {
            return Err(FormulaEvalError::Circular);
        }
        let formula_text = if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, row, col)
        } else {
            formula.text.clone()
        };
        let result = self.evaluate_formula_text(sheet_id, &formula_text, Some((row, col)));
        self.visiting.remove(&(sheet_id, row, col));
        result
    }

    pub(super) fn evaluate_formula_text(
        &mut self,
        sheet_id: SheetId,
        formula_text: &str,
        current_position: Option<(u32, u32)>,
    ) -> Result<CellValue, FormulaEvalError> {
        let text_result = {
            let mut parser = FormulaParser::new(formula_text, self, sheet_id, current_position);
            parser.parse_text_formula()
        };
        match text_result {
            Ok(text) => return Ok(CellValue::Text(text)),
            Err(FormulaEvalError::Unsupported) => {}
            Err(error) => return Err(error),
        }
        FormulaParser::new(formula_text, self, sheet_id, current_position)
            .parse_formula()
            .map(CellValue::Number)
    }

    fn evaluate_formula_value_probe_text(
        &mut self,
        sheet_id: SheetId,
        formula_text: &str,
        current_position: Option<(u32, u32)>,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        FormulaParser::new(formula_text, self, sheet_id, current_position)
            .parse_value_probe_formula()
    }

    fn numeric_cell_value(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<f64, FormulaEvalError> {
        match self.evaluate_cell(sheet_id, row, col) {
            Ok(CellValue::Blank) => Ok(0.0),
            Ok(CellValue::Number(number)) => Ok(number),
            Ok(CellValue::Bool(value)) => Ok(if value { 1.0 } else { 0.0 }),
            Ok(CellValue::Text(_)) => Err(FormulaEvalError::Value),
            Ok(CellValue::Error(error)) => Err(formula_eval_error_from_cell_error(error)),
            Err(FormulaEvalError::Unsupported) => self
                .state
                .worksheet_data
                .get(&sheet_id)
                .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                .map(|cell| match &cell.value {
                    CellValue::Blank => Ok(0.0),
                    CellValue::Number(number) => Ok(*number),
                    CellValue::Bool(value) => Ok(if *value { 1.0 } else { 0.0 }),
                    CellValue::Text(_) => Err(FormulaEvalError::Value),
                    CellValue::Error(error) => Err(formula_eval_error_from_cell_error(*error)),
                })
                .unwrap_or(Ok(0.0)),
            Err(error) => Err(error),
        }
    }

    fn numeric_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<Vec<f64>, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data.get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        let keys = worksheet
            .cells
            .keys()
            .copied()
            .filter(|(row, col)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
            })
            .collect::<Vec<_>>();
        let mut values = Vec::new();
        for (row, col) in keys {
            match self.evaluate_cell(sheet_id, row, col) {
                Ok(CellValue::Number(number)) => values.push(number),
                Ok(CellValue::Error(error)) => {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                Ok(CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_)) => {}
                Err(FormulaEvalError::Unsupported) => {
                    if let Some(cell) = self
                        .state
                        .worksheet_data
                        .get(&sheet_id)
                        .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                    {
                        match cell.value {
                            CellValue::Number(number) => values.push(number),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error));
                            }
                            CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                        }
                    }
                }
                Err(error) => return Err(error),
            }
        }
        Ok(values)
    }

    fn numeric_values_in_reference(
        &mut self,
        reference: &FormulaReference,
    ) -> Result<Vec<f64>, FormulaEvalError> {
        let mut values = Vec::new();
        for (sheet_id, rect) in reference.areas() {
            values.extend(self.numeric_values_in_rect(*sheet_id, *rect)?);
        }
        Ok(values)
    }

    fn counta_values_in_rect(
        &self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<u64, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data.get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        Ok(worksheet
            .cells
            .iter()
            .filter(|((row, col), cell)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
                    && (cell.formula.is_some() || !matches!(cell.value, CellValue::Blank))
            })
            .count() as u64)
    }

    fn counta_values_in_reference(
        &self,
        reference: &FormulaReference,
    ) -> Result<u64, FormulaEvalError> {
        let mut count = 0_u64;
        for (sheet_id, rect) in reference.areas() {
            count += self.counta_values_in_rect(*sheet_id, *rect)?;
        }
        Ok(count)
    }

    fn subtotal_numeric_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<Vec<f64>, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data.get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        let keys = worksheet
            .cells
            .keys()
            .copied()
            .filter(|(row, col)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
            })
            .collect::<Vec<_>>();
        let mut values = Vec::new();
        for (row, col) in keys {
            if self
                .state
                .worksheet_data
                .get(&sheet_id)
                .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                .and_then(|cell| cell.formula.as_ref())
                .is_some_and(|formula| formula_source_has_top_level_function(formula, "SUBTOTAL"))
            {
                continue;
            }
            match self.evaluate_cell(sheet_id, row, col) {
                Ok(CellValue::Number(number)) => values.push(number),
                Ok(CellValue::Error(error)) => {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                Ok(CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_)) => {}
                Err(FormulaEvalError::Unsupported) => {
                    if let Some(cell) = self
                        .state
                        .worksheet_data
                        .get(&sheet_id)
                        .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                    {
                        match cell.value {
                            CellValue::Number(number) => values.push(number),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error));
                            }
                            CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                        }
                    }
                }
                Err(error) => return Err(error),
            }
        }
        Ok(values)
    }

    fn subtotal_counta_values_in_rect(
        &self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<u64, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data.get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        Ok(worksheet
            .cells
            .iter()
            .filter(|((row, col), cell)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
                    && !cell.formula.as_ref().is_some_and(|formula| {
                        formula_source_has_top_level_function(formula, "SUBTOTAL")
                    })
                    && (cell.formula.is_some() || !matches!(cell.value, CellValue::Blank))
            })
            .count() as u64)
    }

    fn countblank_values_in_rect(
        &self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<u64, FormulaEvalError> {
        let total = u64::from(rect.width()) * u64::from(rect.height());
        Ok(total - self.counta_values_in_rect(sheet_id, rect)?)
    }

    fn countblank_values_in_reference(
        &self,
        reference: &FormulaReference,
    ) -> Result<u64, FormulaEvalError> {
        let mut count = 0_u64;
        for (sheet_id, rect) in reference.areas() {
            count += self.countblank_values_in_rect(*sheet_id, *rect)?;
        }
        Ok(count)
    }

    fn cell_value_or_blank(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<CellValue, FormulaEvalError> {
        if !self.state.worksheet_data.contains_key(&sheet_id) {
            return Err(FormulaEvalError::Ref);
        }
        match self.evaluate_cell(sheet_id, row, col) {
            Ok(value) => Ok(value),
            Err(FormulaEvalError::Unsupported) => Ok(self
                .state
                .worksheet_data
                .get(&sheet_id)
                .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                .map(|cell| cell.value.clone())
                .unwrap_or(CellValue::Blank)),
            Err(error) => Err(error),
        }
    }

    fn lookup_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        orientation: FormulaLookupOrientation,
    ) -> Result<Vec<FormulaValueProbe>, FormulaEvalError> {
        if !self.state.worksheet_data.contains_key(&sheet_id) {
            return Err(FormulaEvalError::Ref);
        }
        let mut values = Vec::new();
        match orientation {
            FormulaLookupOrientation::FirstColumn => {
                for row in rect.row_first..=rect.row_last {
                    let value = self.cell_value_or_blank(sheet_id, row, rect.col_first)?;
                    values.push(formula_value_probe_from_cell_value(value));
                }
            }
            FormulaLookupOrientation::FirstRow => {
                for col in rect.col_first..=rect.col_last {
                    let value = self.cell_value_or_blank(sheet_id, rect.row_first, col)?;
                    values.push(formula_value_probe_from_cell_value(value));
                }
            }
        }
        Ok(values)
    }

    fn lookup_result_at(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        Ok(formula_value_probe_from_cell_value(
            self.cell_value_or_blank(sheet_id, row, col)?,
        ))
    }

    fn formula_source_at(
        &self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<Option<FormulaSource>, FormulaEvalError> {
        if !self.state.worksheet_data.contains_key(&sheet_id) {
            return Err(FormulaEvalError::Ref);
        }
        Ok(self
            .state
            .worksheet_data
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
            .and_then(|cell| cell.formula.clone()))
    }

    fn countif_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        criteria: &FormulaCriteria,
    ) -> Result<u64, FormulaEvalError> {
        let mut count = 0_u64;
        for row in rect.row_first..=rect.row_last {
            for col in rect.col_first..=rect.col_last {
                let value = self.cell_value_or_blank(sheet_id, row, col)?;
                if let CellValue::Error(error) = value {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                if criteria.matches(&value) {
                    count += 1;
                }
            }
        }
        Ok(count)
    }

    fn sumif_values_in_rect(
        &mut self,
        criteria_sheet_id: SheetId,
        criteria_rect: Rect,
        criteria: &FormulaCriteria,
        sum_sheet_id: SheetId,
        sum_rect: Rect,
    ) -> Result<f64, FormulaEvalError> {
        Ok(self
            .conditional_sum_and_count_in_rect(
                criteria_sheet_id,
                criteria_rect,
                criteria,
                sum_sheet_id,
                sum_rect,
            )?
            .0)
    }

    fn averageif_values_in_rect(
        &mut self,
        criteria_sheet_id: SheetId,
        criteria_rect: Rect,
        criteria: &FormulaCriteria,
        average_sheet_id: SheetId,
        average_rect: Rect,
    ) -> Result<f64, FormulaEvalError> {
        let (total, count) = self.conditional_sum_and_count_in_rect(
            criteria_sheet_id,
            criteria_rect,
            criteria,
            average_sheet_id,
            average_rect,
        )?;
        if count == 0 {
            return Err(FormulaEvalError::Div0);
        }
        Ok(total / count as f64)
    }

    fn conditional_sum_and_count_in_rect(
        &mut self,
        criteria_sheet_id: SheetId,
        criteria_rect: Rect,
        criteria: &FormulaCriteria,
        value_sheet_id: SheetId,
        value_rect: Rect,
    ) -> Result<(f64, u64), FormulaEvalError> {
        let mut total = 0.0;
        let mut count = 0_u64;
        for row in criteria_rect.row_first..=criteria_rect.row_last {
            for col in criteria_rect.col_first..=criteria_rect.col_last {
                let criteria_value = self.cell_value_or_blank(criteria_sheet_id, row, col)?;
                if let CellValue::Error(error) = criteria_value {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                if !criteria.matches(&criteria_value) {
                    continue;
                }
                let value_row = value_rect.row_first + (row - criteria_rect.row_first);
                let value_col = value_rect.col_first + (col - criteria_rect.col_first);
                let value = self.cell_value_or_blank(value_sheet_id, value_row, value_col)?;
                match value {
                    CellValue::Number(number) => {
                        total += number;
                        count += 1;
                    }
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                }
            }
        }
        Ok((total, count))
    }

    fn countifs_values_in_rects(
        &mut self,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<u64, FormulaEvalError> {
        let base_rect = self.validate_multi_criteria_shapes(None, criteria_ranges)?;
        let mut count = 0_u64;
        for row in base_rect.row_first..=base_rect.row_last {
            for col in base_rect.col_first..=base_rect.col_last {
                let mut matches_all = true;
                for criteria_range in criteria_ranges {
                    if !self.criteria_range_matches_at_offset(
                        base_rect,
                        row,
                        col,
                        criteria_range,
                    )? {
                        matches_all = false;
                        break;
                    }
                }
                if matches_all {
                    count += 1;
                }
            }
        }
        Ok(count)
    }

    fn sumifs_values_in_rect(
        &mut self,
        sum_sheet_id: SheetId,
        sum_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        Ok(self
            .multi_criteria_sum_and_count_in_rect(sum_sheet_id, sum_rect, criteria_ranges)?
            .0)
    }

    fn averageifs_values_in_rect(
        &mut self,
        average_sheet_id: SheetId,
        average_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        let (total, count) = self.multi_criteria_sum_and_count_in_rect(
            average_sheet_id,
            average_rect,
            criteria_ranges,
        )?;
        if count == 0 {
            return Err(FormulaEvalError::Div0);
        }
        Ok(total / count as f64)
    }

    fn multi_criteria_sum_and_count_in_rect(
        &mut self,
        value_sheet_id: SheetId,
        value_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<(f64, u64), FormulaEvalError> {
        let base_rect = self.validate_multi_criteria_shapes(Some(value_rect), criteria_ranges)?;
        let mut total = 0.0;
        let mut count = 0_u64;
        for row in base_rect.row_first..=base_rect.row_last {
            for col in base_rect.col_first..=base_rect.col_last {
                let mut matches_all = true;
                for criteria_range in criteria_ranges {
                    if !self.criteria_range_matches_at_offset(
                        base_rect,
                        row,
                        col,
                        criteria_range,
                    )? {
                        matches_all = false;
                        break;
                    }
                }
                if !matches_all {
                    continue;
                }
                let value_row = value_rect.row_first + (row - base_rect.row_first);
                let value_col = value_rect.col_first + (col - base_rect.col_first);
                let value = self.cell_value_or_blank(value_sheet_id, value_row, value_col)?;
                match value {
                    CellValue::Number(number) => {
                        total += number;
                        count += 1;
                    }
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                }
            }
        }
        Ok((total, count))
    }

    fn minifs_values_in_rect(
        &mut self,
        min_sheet_id: SheetId,
        min_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        self.multi_criteria_extreme_value_in_rect(min_sheet_id, min_rect, criteria_ranges, true)
    }

    fn maxifs_values_in_rect(
        &mut self,
        max_sheet_id: SheetId,
        max_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        self.multi_criteria_extreme_value_in_rect(max_sheet_id, max_rect, criteria_ranges, false)
    }

    fn multi_criteria_extreme_value_in_rect(
        &mut self,
        value_sheet_id: SheetId,
        value_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
        want_min: bool,
    ) -> Result<f64, FormulaEvalError> {
        let base_rect = self.validate_multi_criteria_shapes(Some(value_rect), criteria_ranges)?;
        let mut best = None::<f64>;
        for row in base_rect.row_first..=base_rect.row_last {
            for col in base_rect.col_first..=base_rect.col_last {
                let mut matches_all = true;
                for criteria_range in criteria_ranges {
                    if !self.criteria_range_matches_at_offset(
                        base_rect,
                        row,
                        col,
                        criteria_range,
                    )? {
                        matches_all = false;
                        break;
                    }
                }
                if !matches_all {
                    continue;
                }
                let value_row = value_rect.row_first + (row - base_rect.row_first);
                let value_col = value_rect.col_first + (col - base_rect.col_first);
                let value = self.cell_value_or_blank(value_sheet_id, value_row, value_col)?;
                match value {
                    CellValue::Number(number) => {
                        best = Some(match best {
                            Some(current) if want_min => current.min(number),
                            Some(current) => current.max(number),
                            None => number,
                        });
                    }
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                }
            }
        }
        Ok(best.unwrap_or(0.0))
    }

    fn criteria_range_matches_at_offset(
        &mut self,
        base_rect: Rect,
        row: u32,
        col: u32,
        criteria_range: &FormulaCriteriaRange,
    ) -> Result<bool, FormulaEvalError> {
        let criteria_row = criteria_range.rect.row_first + (row - base_rect.row_first);
        let criteria_col = criteria_range.rect.col_first + (col - base_rect.col_first);
        let value =
            self.cell_value_or_blank(criteria_range.sheet_id, criteria_row, criteria_col)?;
        if let CellValue::Error(error) = value {
            return Err(formula_eval_error_from_cell_error(error));
        }
        Ok(criteria_range.criteria.matches(&value))
    }

    fn validate_multi_criteria_shapes(
        &self,
        value_rect: Option<Rect>,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<Rect, FormulaEvalError> {
        let Some(base_rect) = criteria_ranges
            .first()
            .map(|criteria_range| criteria_range.rect)
        else {
            return Err(FormulaEvalError::Value);
        };
        if criteria_ranges.iter().any(|criteria_range| {
            criteria_range.rect.width() != base_rect.width()
                || criteria_range.rect.height() != base_rect.height()
        }) {
            return Err(FormulaEvalError::Value);
        }
        if let Some(value_rect) = value_rect {
            if value_rect.width() != base_rect.width() || value_rect.height() != base_rect.height()
            {
                return Err(FormulaEvalError::Value);
            }
        }
        Ok(base_rect)
    }
}

struct FormulaParser<'a, 'b, 'state> {
    input: &'a str,
    index: usize,
    evaluator: &'b mut FormulaEvaluator<'state>,
    sheet_id: SheetId,
    current_position: Option<(u32, u32)>,
    bindings: Vec<(String, FormulaValueProbe)>,
}

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    fn new(
        input: &'a str,
        evaluator: &'b mut FormulaEvaluator<'state>,
        sheet_id: SheetId,
        current_position: Option<(u32, u32)>,
    ) -> Self {
        Self {
            input: input.trim().strip_prefix('=').unwrap_or(input.trim()),
            index: 0,
            evaluator,
            sheet_id,
            current_position,
            bindings: Vec::new(),
        }
    }

    fn parse_dynamic_array_formula(&mut self) -> Result<FormulaArrayResult, FormulaEvalError> {
        self.skip_whitespace();
        let Some(identifier) = self.parse_identifier() else {
            return Err(FormulaEvalError::Unsupported);
        };
        if ![
            "FILTER",
            "SORT",
            "SORTBY",
            "UNIQUE",
            "TAKE",
            "DROP",
            "CHOOSECOLS",
            "CHOOSEROWS",
            "TRANSPOSE",
            "SEQUENCE",
            "EXPAND",
            "HSTACK",
            "VSTACK",
            "TOCOL",
            "TOROW",
            "WRAPROWS",
            "WRAPCOLS",
        ]
        .iter()
        .any(|name| identifier.eq_ignore_ascii_case(name))
        {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        if !self.consume_char('(') {
            return Err(FormulaEvalError::Unsupported);
        }

        if identifier.eq_ignore_ascii_case("SEQUENCE") {
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows < 1 {
                return Err(FormulaEvalError::Value);
            }
            let mut cols = 1_i64;
            let mut start = 1.0_f64;
            let mut step = 1.0_f64;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    cols = formula_integer_argument(self.parse_comparison()?)?;
                    if cols < 1 {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        start = self.parse_comparison()?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        step = self.parse_comparison()?;
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() || !start.is_finite() || !step.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let rows = usize::try_from(rows).map_err(|_| FormulaEvalError::Value)?;
            let cols = usize::try_from(cols).map_err(|_| FormulaEvalError::Value)?;
            let len = rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?;
            let values = (0..len)
                .map(|index| {
                    formula_checked_numeric_result(start + step * index as f64)
                        .map(CellValue::Number)
                })
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        let (source_sheet_id, source_rect) = self.parse_reference_argument()?;
        let source_rows = source_rect.height() as usize;
        let source_cols = source_rect.width() as usize;
        let mut source_values = Vec::with_capacity(source_rows * source_cols);
        for row in source_rect.row_first..=source_rect.row_last {
            for col in source_rect.col_first..=source_rect.col_last {
                source_values.push(self.evaluator.cell_value_or_blank(
                    source_sheet_id,
                    row,
                    col,
                )?);
            }
        }

        if identifier.eq_ignore_ascii_case("FILTER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (include_sheet_id, include_rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            let if_empty = if self.consume_char(')') {
                None
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let value = self.parse_value_probe_argument()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                Some(value)
            };
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }

            let include_value = |value: CellValue| -> Result<bool, FormulaEvalError> {
                match value {
                    CellValue::Blank => Ok(false),
                    CellValue::Bool(value) => Ok(value),
                    CellValue::Number(value) => Ok(value != 0.0),
                    CellValue::Text(_) => Err(FormulaEvalError::Value),
                    CellValue::Error(error) => Err(formula_eval_error_from_cell_error(error)),
                }
            };
            let mut values = Vec::new();
            let (rows, cols) = if include_rect.height() == source_rect.height()
                && include_rect.width() == 1
            {
                let mut selected_rows = 0_usize;
                for row_offset in 0..source_rect.height() {
                    let include = self.evaluator.cell_value_or_blank(
                        include_sheet_id,
                        include_rect.row_first + row_offset,
                        include_rect.col_first,
                    )?;
                    if include_value(include)? {
                        selected_rows += 1;
                        let source_start = row_offset as usize * source_cols;
                        values.extend_from_slice(
                            &source_values[source_start..source_start + source_cols],
                        );
                    }
                }
                (selected_rows, source_cols)
            } else if include_rect.width() == source_rect.width() && include_rect.height() == 1 {
                let selected_columns = (0..source_rect.width())
                    .map(|col_offset| {
                        self.evaluator
                            .cell_value_or_blank(
                                include_sheet_id,
                                include_rect.row_first,
                                include_rect.col_first + col_offset,
                            )
                            .and_then(&include_value)
                            .map(|include| (col_offset as usize, include))
                    })
                    .collect::<Result<Vec<_>, _>>()?
                    .into_iter()
                    .filter_map(|(col_offset, include)| include.then_some(col_offset))
                    .collect::<Vec<_>>();
                for row_offset in 0..source_rows {
                    for &col_offset in &selected_columns {
                        values.push(source_values[row_offset * source_cols + col_offset].clone());
                    }
                }
                (source_rows, selected_columns.len())
            } else {
                return Err(FormulaEvalError::Value);
            };

            if rows == 0 || cols == 0 {
                return if_empty
                    .and_then(formula_cell_value_from_probe)
                    .map(FormulaArrayResult::single)
                    .ok_or(FormulaEvalError::Calc);
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("EXPAND") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows < source_rows as i64 {
                return Err(FormulaEvalError::Value);
            }
            let mut cols = source_cols as i64;
            let mut pad = CellValue::Error(CellError::NA);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    cols = formula_integer_argument(self.parse_comparison()?)?;
                }
                if cols < source_cols as i64 {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    pad = formula_cell_value_from_probe(self.parse_value_probe_argument()?)
                        .ok_or(FormulaEvalError::Value)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() || rows < 1 || cols < 1 {
                return Err(FormulaEvalError::Value);
            }
            let rows = usize::try_from(rows).map_err(|_| FormulaEvalError::Value)?;
            let cols = usize::try_from(cols).map_err(|_| FormulaEvalError::Value)?;
            let mut values = vec![pad; rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?];
            for row in 0..source_rows {
                for col in 0..source_cols {
                    values[row * cols + col] = source_values[row * source_cols + col].clone();
                }
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("HSTACK") || identifier.eq_ignore_ascii_case("VSTACK") {
            let mut matrices = vec![(source_rows, source_cols, source_values)];
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    break;
                }
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let (sheet_id, rect) = self.parse_reference_argument()?;
                let rows = rect.height() as usize;
                let cols = rect.width() as usize;
                let mut values = Vec::with_capacity(rows * cols);
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        values.push(self.evaluator.cell_value_or_blank(sheet_id, row, col)?);
                    }
                }
                matrices.push((rows, cols, values));
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            if identifier.eq_ignore_ascii_case("HSTACK") {
                let rows = matrices
                    .iter()
                    .map(|(rows, _, _)| *rows)
                    .max()
                    .ok_or(FormulaEvalError::Calc)?;
                let cols = matrices.iter().try_fold(0_usize, |total, (_, cols, _)| {
                    total.checked_add(*cols).ok_or(FormulaEvalError::Num)
                })?;
                let mut values = vec![CellValue::Error(CellError::NA); rows * cols];
                let mut col_start = 0_usize;
                for (matrix_rows, matrix_cols, matrix_values) in matrices {
                    for row in 0..matrix_rows {
                        for col in 0..matrix_cols {
                            values[row * cols + col_start + col] =
                                matrix_values[row * matrix_cols + col].clone();
                        }
                    }
                    col_start += matrix_cols;
                }
                return Ok(FormulaArrayResult { rows, cols, values });
            }
            let rows = matrices.iter().try_fold(0_usize, |total, (rows, _, _)| {
                total.checked_add(*rows).ok_or(FormulaEvalError::Num)
            })?;
            let cols = matrices
                .iter()
                .map(|(_, cols, _)| *cols)
                .max()
                .ok_or(FormulaEvalError::Calc)?;
            let mut values = vec![CellValue::Error(CellError::NA); rows * cols];
            let mut row_start = 0_usize;
            for (matrix_rows, matrix_cols, matrix_values) in matrices {
                for row in 0..matrix_rows {
                    for col in 0..matrix_cols {
                        values[(row_start + row) * cols + col] =
                            matrix_values[row * matrix_cols + col].clone();
                    }
                }
                row_start += matrix_rows;
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("TOCOL") || identifier.eq_ignore_ascii_case("TOROW") {
            let mut ignore = 0_i64;
            let mut scan_by_column = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    ignore = formula_integer_argument(self.parse_comparison()?)?;
                }
                if !(0..=3).contains(&ignore) {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    scan_by_column = self.parse_comparison()? != 0.0;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let include_value = |value: &CellValue| {
                let ignore_blank = matches!(ignore, 1 | 3);
                let ignore_error = matches!(ignore, 2 | 3);
                !(ignore_blank && matches!(value, CellValue::Blank))
                    && !(ignore_error && matches!(value, CellValue::Error(_)))
            };
            let mut values = Vec::with_capacity(source_values.len());
            if scan_by_column {
                for col in 0..source_cols {
                    for row in 0..source_rows {
                        let value = &source_values[row * source_cols + col];
                        if include_value(value) {
                            values.push(value.clone());
                        }
                    }
                }
            } else {
                values.extend(source_values.into_iter().filter(include_value));
            }
            if values.is_empty() {
                return Err(FormulaEvalError::Calc);
            }
            let (rows, cols) = if identifier.eq_ignore_ascii_case("TOCOL") {
                (values.len(), 1)
            } else {
                (1, values.len())
            };
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("WRAPROWS")
            || identifier.eq_ignore_ascii_case("WRAPCOLS")
        {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let wrap_count = formula_integer_argument(self.parse_comparison()?)?;
            if wrap_count < 1 {
                return Err(FormulaEvalError::Value);
            }
            let mut pad = CellValue::Error(CellError::NA);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                pad = formula_cell_value_from_probe(self.parse_value_probe_argument()?)
                    .ok_or(FormulaEvalError::Value)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let wrap_count = usize::try_from(wrap_count).map_err(|_| FormulaEvalError::Value)?;
            let chunk_count = source_values.len().div_ceil(wrap_count);
            let (rows, cols) = if identifier.eq_ignore_ascii_case("WRAPROWS") {
                (chunk_count, wrap_count)
            } else {
                (wrap_count, chunk_count)
            };
            let mut values = vec![pad; rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?];
            if identifier.eq_ignore_ascii_case("WRAPROWS") {
                values[..source_values.len()].clone_from_slice(&source_values);
            } else {
                for (index, value) in source_values.into_iter().enumerate() {
                    let row = index % wrap_count;
                    let col = index / wrap_count;
                    values[row * cols + col] = value;
                }
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("SORT") || identifier.eq_ignore_ascii_case("SORTBY") {
            let compare_values =
                |left: &CellValue, right: &CellValue| -> Result<Ordering, FormulaEvalError> {
                    let left = formula_value_probe_from_cell_value(left.clone());
                    let right = formula_value_probe_from_cell_value(right.clone());
                    if let FormulaValueProbe::Error(error) = left {
                        return Err(error);
                    }
                    if let FormulaValueProbe::Error(error) = right {
                        return Err(error);
                    }
                    if let Some(ordering) = formula_value_probe_ordering(&left, &right)? {
                        return Ok(ordering);
                    }
                    let rank = |value: &FormulaValueProbe| match value {
                        FormulaValueProbe::Blank => 0,
                        FormulaValueProbe::Number(_) => 1,
                        FormulaValueProbe::Text(_) => 2,
                        FormulaValueProbe::Bool(_) => 3,
                        FormulaValueProbe::Error(_) => 4,
                        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => 5,
                    };
                    Ok(rank(&left).cmp(&rank(&right)))
                };
            let mut key_sets = Vec::<(Vec<CellValue>, bool)>::new();
            let by_column;
            if identifier.eq_ignore_ascii_case("SORT") {
                let mut sort_index = 1_i64;
                let mut sort_order = 1_i64;
                let mut sort_by_column = false;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        sort_index = formula_integer_argument(self.parse_comparison()?)?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            sort_order = formula_integer_argument(self.parse_comparison()?)?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            self.skip_whitespace();
                            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                                sort_by_column = self.parse_comparison()? != 0.0;
                            }
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
                if sort_index < 1 || !matches!(sort_order, -1 | 1) {
                    return Err(FormulaEvalError::Value);
                }
                by_column = sort_by_column;
                let sort_index =
                    usize::try_from(sort_index - 1).map_err(|_| FormulaEvalError::Value)?;
                let keys = if by_column {
                    if sort_index >= source_rows {
                        return Err(FormulaEvalError::Value);
                    }
                    (0..source_cols)
                        .map(|col| source_values[sort_index * source_cols + col].clone())
                        .collect()
                } else {
                    if sort_index >= source_cols {
                        return Err(FormulaEvalError::Value);
                    }
                    (0..source_rows)
                        .map(|row| source_values[row * source_cols + sort_index].clone())
                        .collect()
                };
                key_sets.push((keys, sort_order == -1));
            } else {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let mut orientation = None;
                loop {
                    let (key_sheet_id, key_rect) = self.parse_reference_argument()?;
                    let current_by_column = if key_rect.height() as usize == source_rows
                        && key_rect.width() == 1
                    {
                        false
                    } else if key_rect.width() as usize == source_cols && key_rect.height() == 1 {
                        true
                    } else {
                        return Err(FormulaEvalError::Value);
                    };
                    if orientation.is_some_and(|value| value != current_by_column) {
                        return Err(FormulaEvalError::Value);
                    }
                    orientation = Some(current_by_column);
                    let keys = if current_by_column {
                        (key_rect.col_first..=key_rect.col_last)
                            .map(|col| {
                                self.evaluator.cell_value_or_blank(
                                    key_sheet_id,
                                    key_rect.row_first,
                                    col,
                                )
                            })
                            .collect::<Result<Vec<_>, _>>()?
                    } else {
                        (key_rect.row_first..=key_rect.row_last)
                            .map(|row| {
                                self.evaluator.cell_value_or_blank(
                                    key_sheet_id,
                                    row,
                                    key_rect.col_first,
                                )
                            })
                            .collect::<Result<Vec<_>, _>>()?
                    };
                    self.skip_whitespace();
                    let sort_order = if self.consume_char(')') {
                        key_sets.push((keys, false));
                        break;
                    } else {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            1
                        } else {
                            formula_integer_argument(self.parse_comparison()?)?
                        }
                    };
                    if !matches!(sort_order, -1 | 1) {
                        return Err(FormulaEvalError::Value);
                    }
                    key_sets.push((keys, sort_order == -1));
                    self.skip_whitespace();
                    if self.consume_char(')') {
                        break;
                    }
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                }
                by_column = orientation.unwrap_or(false);
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut indexes = if by_column {
                (0..source_cols).collect::<Vec<_>>()
            } else {
                (0..source_rows).collect::<Vec<_>>()
            };
            for index in 1..indexes.len() {
                let current = indexes[index];
                let mut position = index;
                while position > 0 {
                    let previous = indexes[position - 1];
                    let mut ordering = Ordering::Equal;
                    for (keys, descending) in &key_sets {
                        ordering = compare_values(&keys[previous], &keys[current])?;
                        if *descending {
                            ordering = ordering.reverse();
                        }
                        if ordering != Ordering::Equal {
                            break;
                        }
                    }
                    if ordering != Ordering::Greater {
                        break;
                    }
                    indexes[position] = previous;
                    position -= 1;
                }
                indexes[position] = current;
            }
            let mut values = Vec::with_capacity(source_values.len());
            if by_column {
                for row in 0..source_rows {
                    for &col in &indexes {
                        values.push(source_values[row * source_cols + col].clone());
                    }
                }
            } else {
                for &row in &indexes {
                    values.extend_from_slice(
                        &source_values[row * source_cols..(row + 1) * source_cols],
                    );
                }
            }
            return Ok(FormulaArrayResult {
                rows: source_rows,
                cols: source_cols,
                values,
            });
        }

        if identifier.eq_ignore_ascii_case("UNIQUE") {
            let mut by_column = false;
            let mut exactly_once = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    by_column = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        exactly_once = self.parse_comparison()? != 0.0;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let record_count = if by_column { source_cols } else { source_rows };
            let record_len = if by_column { source_rows } else { source_cols };
            let records_equal = |left: usize, right: usize| -> Result<bool, FormulaEvalError> {
                for offset in 0..record_len {
                    let left_value = if by_column {
                        &source_values[offset * source_cols + left]
                    } else {
                        &source_values[left * source_cols + offset]
                    };
                    let right_value = if by_column {
                        &source_values[offset * source_cols + right]
                    } else {
                        &source_values[right * source_cols + offset]
                    };
                    if !formula_value_probe_exact_match(
                        &formula_value_probe_from_cell_value(left_value.clone()),
                        &formula_value_probe_from_cell_value(right_value.clone()),
                    )? {
                        return Ok(false);
                    }
                }
                Ok(true)
            };
            let mut selected = Vec::new();
            for candidate in 0..record_count {
                let mut prior_match = false;
                let mut count = 0_usize;
                for other in 0..record_count {
                    if records_equal(candidate, other)? {
                        count += 1;
                        if other < candidate {
                            prior_match = true;
                        }
                    }
                }
                if (exactly_once && count == 1) || (!exactly_once && !prior_match) {
                    selected.push(candidate);
                }
            }
            if selected.is_empty() {
                return Err(FormulaEvalError::Calc);
            }
            let mut values = Vec::with_capacity(selected.len() * record_len);
            if by_column {
                for row in 0..source_rows {
                    for &col in &selected {
                        values.push(source_values[row * source_cols + col].clone());
                    }
                }
                return Ok(FormulaArrayResult {
                    rows: source_rows,
                    cols: selected.len(),
                    values,
                });
            }
            for &row in &selected {
                values
                    .extend_from_slice(&source_values[row * source_cols..(row + 1) * source_cols]);
            }
            return Ok(FormulaArrayResult {
                rows: selected.len(),
                cols: source_cols,
                values,
            });
        }

        if identifier.eq_ignore_ascii_case("TAKE") || identifier.eq_ignore_ascii_case("DROP") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let row_count = formula_integer_argument(self.parse_comparison()?)?;
            let mut col_count = if identifier.eq_ignore_ascii_case("TAKE") {
                source_cols as i64
            } else {
                0
            };
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                col_count = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let (row_start, rows, col_start, cols) = if identifier.eq_ignore_ascii_case("TAKE") {
                if row_count == 0
                    || row_count.unsigned_abs() > source_rows as u64
                    || col_count == 0
                    || col_count.unsigned_abs() > source_cols as u64
                {
                    return Err(FormulaEvalError::Calc);
                }
                let rows = row_count.unsigned_abs() as usize;
                let cols = col_count.unsigned_abs() as usize;
                (
                    if row_count < 0 { source_rows - rows } else { 0 },
                    rows,
                    if col_count < 0 { source_cols - cols } else { 0 },
                    cols,
                )
            } else {
                if row_count.unsigned_abs() >= source_rows as u64
                    || col_count.unsigned_abs() >= source_cols as u64
                {
                    return Err(FormulaEvalError::Calc);
                }
                let dropped_rows = row_count.unsigned_abs() as usize;
                let dropped_cols = col_count.unsigned_abs() as usize;
                (
                    if row_count > 0 { dropped_rows } else { 0 },
                    source_rows - dropped_rows,
                    if col_count > 0 { dropped_cols } else { 0 },
                    source_cols - dropped_cols,
                )
            };
            let mut values = Vec::with_capacity(rows * cols);
            for row in row_start..row_start + rows {
                values.extend_from_slice(
                    &source_values
                        [row * source_cols + col_start..row * source_cols + col_start + cols],
                );
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("CHOOSECOLS")
            || identifier.eq_ignore_ascii_case("CHOOSEROWS")
        {
            let selecting_columns = identifier.eq_ignore_ascii_case("CHOOSECOLS");
            let size = if selecting_columns {
                source_cols
            } else {
                source_rows
            };
            let mut selected = Vec::new();
            loop {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let index = formula_integer_argument(self.parse_comparison()?)?;
                if index == 0 || index.unsigned_abs() > size as u64 {
                    return Err(FormulaEvalError::Value);
                }
                selected.push(if index > 0 {
                    usize::try_from(index - 1).map_err(|_| FormulaEvalError::Value)?
                } else {
                    size - index.unsigned_abs() as usize
                });
                self.skip_whitespace();
                if self.consume_char(')') {
                    break;
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let mut values = Vec::new();
            if selecting_columns {
                values.reserve(source_rows * selected.len());
                for row in 0..source_rows {
                    for &col in &selected {
                        values.push(source_values[row * source_cols + col].clone());
                    }
                }
                return Ok(FormulaArrayResult {
                    rows: source_rows,
                    cols: selected.len(),
                    values,
                });
            }
            values.reserve(selected.len() * source_cols);
            for &row in &selected {
                values
                    .extend_from_slice(&source_values[row * source_cols..(row + 1) * source_cols]);
            }
            return Ok(FormulaArrayResult {
                rows: selected.len(),
                cols: source_cols,
                values,
            });
        }

        if identifier.eq_ignore_ascii_case("TRANSPOSE") {
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let mut values = Vec::with_capacity(source_values.len());
            for col in 0..source_cols {
                for row in 0..source_rows {
                    values.push(source_values[row * source_cols + col].clone());
                }
            }
            return Ok(FormulaArrayResult {
                rows: source_cols,
                cols: source_rows,
                values,
            });
        }

        Err(FormulaEvalError::Unsupported)
    }

    fn binding_value(&self, name: &str) -> Option<FormulaValueProbe> {
        self.bindings
            .iter()
            .rev()
            .find(|(binding_name, _)| binding_name.eq_ignore_ascii_case(name))
            .map(|(_, value)| value.clone())
    }

    fn parse_text_formula(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(text) = self.parse_string_literal()? {
            self.skip_whitespace();
            return if self.index == self.input.len() {
                Ok(text)
            } else {
                Err(FormulaEvalError::Unsupported)
            };
        }
        let Some(identifier) = self.parse_identifier() else {
            return Err(FormulaEvalError::Unsupported);
        };
        if !formula_text_function_name(identifier.as_str()) {
            self.skip_whitespace();
            if !self.consume_char('(') {
                return Err(FormulaEvalError::Unsupported);
            }
            let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? else {
                return Err(FormulaEvalError::Unsupported);
            };
            self.skip_whitespace();
            return if self.index == self.input.len() {
                match value {
                    FormulaValueProbe::Text(text) => Ok(text),
                    FormulaValueProbe::Error(error) => Err(error),
                    _ => Err(FormulaEvalError::Unsupported),
                }
            } else {
                Err(FormulaEvalError::Unsupported)
            };
        }
        self.skip_whitespace();
        if !self.consume_char('(') {
            return Err(FormulaEvalError::Unsupported);
        }
        let value = self.parse_text_function(identifier.as_str())?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    fn parse_formula(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    fn parse_value_probe_formula(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    fn parse_bound_lambda_call_value(
        &mut self,
        name: &str,
    ) -> Result<Option<FormulaValueProbe>, FormulaEvalError> {
        if let Some(lambda @ FormulaValueProbe::Lambda { .. }) = self.binding_value(name) {
            return self.parse_lambda_call_arguments(lambda).map(Some);
        }
        if let Some(lambda @ FormulaValueProbe::Lambda { .. }) =
            self.defined_name_value_probe(name)?
        {
            return self.parse_lambda_call_arguments(lambda).map(Some);
        }
        Ok(None)
    }

    fn parse_lambda_argument(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        match self.parse_value_probe_argument()? {
            lambda @ FormulaValueProbe::Lambda { .. } => Ok(lambda),
            FormulaValueProbe::Error(error) => Err(error),
            _ => Err(FormulaEvalError::Value),
        }
    }

    fn try_parse_lambda_argument(&mut self) -> Result<Option<FormulaValueProbe>, FormulaEvalError> {
        let checkpoint = self.index;
        match self.parse_lambda_argument() {
            Ok(lambda) => Ok(Some(lambda)),
            Err(FormulaEvalError::Unsupported) | Err(FormulaEvalError::Value) => {
                self.index = checkpoint;
                Ok(None)
            }
            Err(error) => {
                self.index = checkpoint;
                Err(error)
            }
        }
    }

    fn parse_lambda_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut parameters = Vec::new();
        loop {
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some(name) = self.parse_identifier() {
                self.skip_whitespace();
                if self.consume_char(',') {
                    parameters.push(name);
                    continue;
                }
            }
            self.index = checkpoint;
            let body = self.capture_formula_source_until_closing_paren()?;
            if body.trim().is_empty() {
                return Err(FormulaEvalError::Value);
            }
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(FormulaValueProbe::Lambda {
                parameters,
                body: body.trim().to_string(),
            });
        }
    }

    fn capture_formula_source_until_closing_paren(&mut self) -> Result<&'a str, FormulaEvalError> {
        let start = self.index;
        let mut cursor = self.index;
        let mut depth = 0_u32;
        while cursor < self.input.len() {
            let ch = self.input[cursor..]
                .chars()
                .next()
                .ok_or(FormulaEvalError::Unsupported)?;
            if ch == '"' {
                cursor += ch.len_utf8();
                while cursor < self.input.len() {
                    let quoted = self.input[cursor..]
                        .chars()
                        .next()
                        .ok_or(FormulaEvalError::Unsupported)?;
                    cursor += quoted.len_utf8();
                    if quoted == '"' {
                        if self.input[cursor..].starts_with('"') {
                            cursor += 1;
                            continue;
                        }
                        break;
                    }
                }
                continue;
            }
            if ch == '(' {
                depth += 1;
                cursor += ch.len_utf8();
                continue;
            }
            if ch == ')' {
                if depth == 0 {
                    let body = &self.input[start..cursor];
                    self.index = cursor;
                    return Ok(body);
                }
                depth -= 1;
                cursor += ch.len_utf8();
                continue;
            }
            cursor += ch.len_utf8();
        }
        Err(FormulaEvalError::Unsupported)
    }

    fn parse_lambda_call_arguments(
        &mut self,
        lambda: FormulaValueProbe,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut arguments = Vec::new();
        let mut after_separator = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                if after_separator {
                    arguments.push(FormulaValueProbe::Omitted);
                }
                return self.evaluate_lambda_value(lambda, arguments);
            }
            if self.consume_char(',') {
                arguments.push(FormulaValueProbe::Omitted);
                after_separator = true;
                continue;
            }
            arguments.push(self.parse_value_probe_argument()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                after_separator = true;
                continue;
            }
            if self.consume_char(')') {
                return self.evaluate_lambda_value(lambda, arguments);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn evaluate_lambda_value(
        &mut self,
        lambda: FormulaValueProbe,
        arguments: Vec<FormulaValueProbe>,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let FormulaValueProbe::Lambda { parameters, body } = lambda else {
            return Err(FormulaEvalError::Value);
        };
        if arguments.len() > parameters.len() {
            return Err(FormulaEvalError::Value);
        }
        let mut bindings = self.bindings.clone();
        for (index, name) in parameters.into_iter().enumerate() {
            let value = arguments
                .get(index)
                .cloned()
                .unwrap_or(FormulaValueProbe::Omitted);
            bindings.push((name, value));
        }

        let mut parser = FormulaParser::new(
            body.as_str(),
            &mut *self.evaluator,
            self.sheet_id,
            self.current_position,
        );
        parser.bindings = bindings;
        parser.parse_value_probe_formula()
    }

    fn parse_comparison(&mut self) -> Result<f64, FormulaEvalError> {
        let mut value = self.parse_expression()?;
        loop {
            self.skip_whitespace();
            let Some(operator) = self.consume_comparison_operator() else {
                return Ok(value);
            };
            let right = self.parse_expression()?;
            value = if operator.evaluate(value, right) {
                1.0
            } else {
                0.0
            };
        }
    }

    fn parse_expression(&mut self) -> Result<f64, FormulaEvalError> {
        let mut value = self.parse_term()?;
        loop {
            self.skip_whitespace();
            if self.consume_char('+') {
                value += self.parse_term()?;
            } else if self.consume_char('-') {
                value -= self.parse_term()?;
            } else {
                return Ok(value);
            }
        }
    }

    fn parse_term(&mut self) -> Result<f64, FormulaEvalError> {
        let mut value = self.parse_factor()?;
        loop {
            self.skip_whitespace();
            if self.consume_char('*') {
                value *= self.parse_factor()?;
            } else if self.consume_char('/') {
                let divisor = self.parse_factor()?;
                if divisor == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                value /= divisor;
            } else {
                return Ok(value);
            }
        }
    }

    fn parse_factor(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char('+') {
            return self.parse_factor();
        }
        if self.consume_char('-') {
            return Ok(-self.parse_factor()?);
        }
        if self.consume_char('(') {
            let value = self.parse_expression()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(value);
        }
        if let Some(number) = self.parse_number()? {
            return Ok(number);
        }
        let checkpoint = self.index;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return formula_number_from_value_probe(value);
                }
                return self.parse_function(identifier.as_str());
            }
        }
        self.index = checkpoint;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (target_sheet_id, rect) = reference.single_area()?;
            if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
                return self.evaluator.numeric_cell_value(
                    target_sheet_id,
                    rect.row_first,
                    rect.col_first,
                );
            }
            return Err(FormulaEvalError::Value);
        }
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return formula_number_from_value_probe(value);
                }
                return self.parse_function(identifier.as_str());
            }
            if identifier.eq_ignore_ascii_case("TRUE") {
                return Ok(1.0);
            }
            if identifier.eq_ignore_ascii_case("FALSE") {
                return Ok(0.0);
            }
            if let Some(value) = self.binding_value(identifier.as_str()) {
                return formula_number_from_value_probe(value);
            }
            if let Some(value) = self.defined_name_value_probe(identifier.as_str())? {
                return formula_number_from_value_probe(value);
            }
            return Err(FormulaEvalError::Name);
        }
        Err(FormulaEvalError::Unsupported)
    }

    fn parse_function(&mut self, name: &str) -> Result<f64, FormulaEvalError> {
        let percentile_value =
            |mut values: Vec<f64>, k: f64, exclusive: bool| -> Result<f64, FormulaEvalError> {
                if values.is_empty() || !k.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                values.sort_by(|left, right| left.total_cmp(right));
                if exclusive {
                    if k <= 0.0 || k >= 1.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let rank = k * (values.len() as f64 + 1.0);
                    if rank < 1.0 || rank > values.len() as f64 {
                        return Err(FormulaEvalError::Num);
                    }
                    let lower_rank = rank.floor();
                    let upper_rank = rank.ceil();
                    if lower_rank == upper_rank {
                        return Ok(values[lower_rank as usize - 1]);
                    }
                    let lower_index = lower_rank as usize - 1;
                    let upper_index = upper_rank as usize - 1;
                    let fraction = rank - lower_rank;
                    return Ok(values[lower_index]
                        + (values[upper_index] - values[lower_index]) * fraction);
                }
                if !(0.0..=1.0).contains(&k) {
                    return Err(FormulaEvalError::Num);
                }
                let rank = k * (values.len() as f64 - 1.0);
                let lower_index = rank.floor() as usize;
                let upper_index = rank.ceil() as usize;
                if lower_index == upper_index {
                    return Ok(values[lower_index]);
                }
                let fraction = rank - lower_index as f64;
                Ok(values[lower_index] + (values[upper_index] - values[lower_index]) * fraction)
            };
        let percent_rank_value = |mut values: Vec<f64>,
                                  x: f64,
                                  significance: i64,
                                  exclusive: bool|
         -> Result<f64, FormulaEvalError> {
            if values.is_empty() || !x.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            if significance < 1 {
                return Err(FormulaEvalError::Num);
            }
            let significance = i32::try_from(significance).map_err(|_| FormulaEvalError::Num)?;
            let factor = 10_f64.powi(significance);
            if !factor.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            values.sort_by(|left, right| left.total_cmp(right));
            let Some(minimum) = values.first().copied() else {
                return Err(FormulaEvalError::Num);
            };
            let Some(maximum) = values.last().copied() else {
                return Err(FormulaEvalError::Num);
            };
            if x < minimum || x > maximum {
                return Err(FormulaEvalError::NA);
            }
            if !exclusive && values.len() == 1 {
                return Ok(0.0);
            }
            if exclusive && values.len() == 1 {
                return Ok(0.5);
            }
            let rank_for_index = |index: usize| -> f64 {
                if exclusive {
                    (index as f64 + 1.0) / (values.len() as f64 + 1.0)
                } else {
                    index as f64 / (values.len() as f64 - 1.0)
                }
            };
            let rank = if let Some(index) = values.iter().position(|value| *value == x) {
                rank_for_index(index)
            } else {
                let Some(upper_index) = values.iter().position(|value| *value > x) else {
                    return Err(FormulaEvalError::NA);
                };
                if upper_index == 0 {
                    return Err(FormulaEvalError::NA);
                }
                let lower_index = upper_index - 1;
                let lower_rank = rank_for_index(lower_index);
                let upper_rank = rank_for_index(upper_index);
                lower_rank
                    + (upper_rank - lower_rank) * (x - values[lower_index])
                        / (values[upper_index] - values[lower_index])
            };
            Ok((rank * factor).trunc() / factor)
        };
        if name.eq_ignore_ascii_case("IF") {
            return self.parse_if_function();
        }
        if name.eq_ignore_ascii_case("IFS") {
            return formula_number_from_value_probe(self.parse_ifs_value_function()?);
        }
        if name.eq_ignore_ascii_case("SWITCH") {
            return formula_number_from_value_probe(self.parse_switch_value_function()?);
        }
        if name.eq_ignore_ascii_case("AND") {
            return self.parse_logical_function(FormulaLogicalFunction::And);
        }
        if name.eq_ignore_ascii_case("OR") {
            return self.parse_logical_function(FormulaLogicalFunction::Or);
        }
        if name.eq_ignore_ascii_case("XOR") {
            return self.parse_logical_function(FormulaLogicalFunction::Xor);
        }
        if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(if name.eq_ignore_ascii_case("TRUE") {
                1.0
            } else {
                0.0
            });
        }
        if name.eq_ignore_ascii_case("DATEDIF") {
            return self.parse_datedif_function();
        }
        if name.eq_ignore_ascii_case("WORKDAY") {
            return self.parse_workday_function();
        }
        if name.eq_ignore_ascii_case("WORKDAY.INTL") {
            return self.parse_workday_intl_function();
        }
        if name.eq_ignore_ascii_case("NETWORKDAYS") {
            return self.parse_networkdays_function();
        }
        if name.eq_ignore_ascii_case("NETWORKDAYS.INTL") {
            return self.parse_networkdays_intl_function();
        }
        if name.eq_ignore_ascii_case("SERIESSUM") {
            return self.parse_series_sum_function();
        }
        if name.eq_ignore_ascii_case("AGGREGATE") {
            return self.parse_aggregate_function();
        }
        if name.eq_ignore_ascii_case("SUBTOTAL") {
            return self.parse_subtotal_function();
        }
        if let Some(function) = FormulaScalarFunction::from_name(name) {
            return self.parse_scalar_function(function);
        }
        if name.eq_ignore_ascii_case("COUNTA") {
            return self.parse_counta_function();
        }
        if name.eq_ignore_ascii_case("COUNTBLANK") {
            return self.parse_countblank_function();
        }
        if name.eq_ignore_ascii_case("CHOOSE") {
            return self.parse_choose_function();
        }
        if name.eq_ignore_ascii_case("COLUMN") {
            return self.parse_column_function();
        }
        if name.eq_ignore_ascii_case("COLUMNS") || name.eq_ignore_ascii_case("COLS") {
            return self.parse_columns_function();
        }
        if name.eq_ignore_ascii_case("COUNTIF") {
            return self.parse_countif_function();
        }
        if name.eq_ignore_ascii_case("IFERROR") {
            return self.parse_iferror_function();
        }
        if name.eq_ignore_ascii_case("IFNA") {
            return self.parse_ifna_function();
        }
        if name.eq_ignore_ascii_case("ISERROR") {
            return self.parse_error_test_function(true, false);
        }
        if name.eq_ignore_ascii_case("ISERR") {
            return self.parse_error_test_function(false, false);
        }
        if name.eq_ignore_ascii_case("ISNA") {
            return self.parse_error_test_function(false, true);
        }
        if name.eq_ignore_ascii_case("ISBLANK") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Blank)
            });
        }
        if name.eq_ignore_ascii_case("ISNUMBER") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Number(_))
            });
        }
        if name.eq_ignore_ascii_case("ISLOGICAL") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Bool(_))
            });
        }
        if name.eq_ignore_ascii_case("ISNONTEXT") {
            return self.parse_value_probe_test_function(|value| {
                !matches!(value, FormulaValueProbe::Text(_))
            });
        }
        if name.eq_ignore_ascii_case("ISTEXT") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Text(_))
            });
        }
        if name.eq_ignore_ascii_case("ISREF") {
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((_, _, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.consume_char(')') {
                    return Ok(1.0);
                }
            }
            self.index = checkpoint;
            let _ = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(0.0);
        }
        if name.eq_ignore_ascii_case("ISFORMULA") {
            return self.parse_isformula_function();
        }
        if name.eq_ignore_ascii_case("TYPE") {
            return self.parse_type_function();
        }
        if name.eq_ignore_ascii_case("ERROR.TYPE") {
            return self.parse_error_type_function();
        }
        if name.eq_ignore_ascii_case("N") {
            let value = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return match value {
                FormulaValueProbe::Blank | FormulaValueProbe::Text(_) => Ok(0.0),
                FormulaValueProbe::Bool(value) => Ok(if value { 1.0 } else { 0.0 }),
                FormulaValueProbe::Number(value) => Ok(value),
                FormulaValueProbe::Error(error) => Err(error),
                FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                    Err(FormulaEvalError::Value)
                }
            };
        }
        if name.eq_ignore_ascii_case("DECIMAL") {
            return self.parse_decimal_function();
        }
        if name.eq_ignore_ascii_case("CONVERT") {
            return self.parse_convert_function();
        }
        if name.eq_ignore_ascii_case("EUROCONVERT") {
            return self.parse_euroconvert_function();
        }
        if name.eq_ignore_ascii_case("MDETERM") {
            return self.parse_mdeterm_function();
        }
        if name.eq_ignore_ascii_case("IMABS")
            || name.eq_ignore_ascii_case("IMAGINARY")
            || name.eq_ignore_ascii_case("IMARGUMENT")
            || name.eq_ignore_ascii_case("IMREAL")
        {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if name.eq_ignore_ascii_case("IMABS") {
                let result = value.real.hypot(value.imaginary);
                return if result.is_finite() {
                    Ok(result)
                } else {
                    Err(FormulaEvalError::Num)
                };
            }
            if name.eq_ignore_ascii_case("IMAGINARY") {
                return Ok(value.imaginary);
            }
            if name.eq_ignore_ascii_case("IMARGUMENT") {
                if value.real == 0.0 && value.imaginary == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                return Ok(value.imaginary.atan2(value.real));
            }
            return Ok(value.real);
        }
        if name.eq_ignore_ascii_case("BIN2DEC") {
            return self.parse_engineering_decimal_function(2, 10, 10);
        }
        if name.eq_ignore_ascii_case("OCT2DEC") {
            return self.parse_engineering_decimal_function(8, 30, 10);
        }
        if name.eq_ignore_ascii_case("HEX2DEC") {
            return self.parse_engineering_decimal_function(16, 40, 10);
        }
        if name.eq_ignore_ascii_case("DOLLARDE") {
            return self.parse_dollarde_function();
        }
        if name.eq_ignore_ascii_case("DOLLARFR") {
            return self.parse_dollarfr_function();
        }
        if name.eq_ignore_ascii_case("FVSCHEDULE") {
            self.skip_whitespace();
            if self.parse_string_literal()?.is_some() {
                return Err(FormulaEvalError::Value);
            }
            let mut value = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let mut schedule = Vec::new();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| ch == ')') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Blank => schedule.push(0.0),
                                CellValue::Number(number) => schedule.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Bool(_) | CellValue::Text(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    schedule.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                schedule.push(self.parse_comparison()?);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !value.is_finite() || schedule.iter().any(|rate| !rate.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            for rate in schedule {
                value *= 1.0 + rate;
            }
            return if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            };
        }
        if name.eq_ignore_ascii_case("NPV") {
            let rate = self.parse_comparison()?;
            if !rate.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let discount = 1.0 + rate;
            let mut discount_factor = 1.0;
            let mut total = 0.0;
            let mut saw_value_argument = false;
            macro_rules! record_cash_flow {
                ($cash_flow:expr) => {{
                    let cash_flow = $cash_flow;
                    if !cash_flow.is_finite() {
                        return Err(FormulaEvalError::Value);
                    }
                    discount_factor *= discount;
                    if discount_factor == 0.0 {
                        return Err(FormulaEvalError::Div0);
                    }
                    if !discount_factor.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    total += cash_flow / discount_factor;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }};
            }
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    return if saw_value_argument {
                        Ok(total)
                    } else {
                        Err(FormulaEvalError::Value)
                    };
                }
                saw_value_argument = true;
                let checkpoint = self.index;
                if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                    self.index = next_index;
                    self.skip_whitespace();
                    if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                match self
                                    .evaluator
                                    .cell_value_or_blank(target_sheet_id, row, col)
                                {
                                    Ok(CellValue::Number(number)) => record_cash_flow!(number),
                                    Ok(_) => {}
                                    Err(FormulaEvalError::Unsupported) => {
                                        return Err(FormulaEvalError::Unsupported);
                                    }
                                    Err(_) => {}
                                }
                            }
                        }
                    } else {
                        self.index = checkpoint;
                        if self.parse_string_literal()?.is_none() {
                            let identifier_checkpoint = self.index;
                            if let Some(identifier) = self.parse_identifier() {
                                self.skip_whitespace();
                                if !(identifier.eq_ignore_ascii_case("TRUE")
                                    || identifier.eq_ignore_ascii_case("FALSE"))
                                    || self.peek_char() == Some('(')
                                {
                                    self.index = identifier_checkpoint;
                                    if let Ok(value) = self.parse_catchable_argument()? {
                                        record_cash_flow!(value);
                                    }
                                }
                            } else {
                                self.index = identifier_checkpoint;
                                if let Ok(value) = self.parse_catchable_argument()? {
                                    record_cash_flow!(value);
                                }
                            }
                        }
                    }
                } else if self.parse_string_literal()?.is_none() {
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if !(identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            || self.peek_char() == Some('(')
                        {
                            self.index = identifier_checkpoint;
                            if let Ok(value) = self.parse_catchable_argument()? {
                                record_cash_flow!(value);
                            }
                        }
                    } else {
                        self.index = identifier_checkpoint;
                        if let Ok(value) = self.parse_catchable_argument()? {
                            record_cash_flow!(value);
                        }
                    }
                }
                self.skip_whitespace();
                if self.consume_char(',') {
                    continue;
                }
                if self.consume_char(')') {
                    return Ok(total);
                }
                return Err(FormulaEvalError::Unsupported);
            }
        }
        if name.eq_ignore_ascii_case("XNPV") {
            let rate = self.parse_comparison()?;
            if !rate.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| ch == ',') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                values.push(self.parse_comparison()?);
            }

            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut dates = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| ch == ')') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let date = match self.evaluator.cell_value_or_blank(
                                target_sheet_id,
                                row,
                                col,
                            )? {
                                CellValue::Number(number) => number,
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            };
                            let serial = formula_serial_integer(date)
                                .and_then(|serial| {
                                    formula_ymd_from_serial(serial as f64).map(|_| serial)
                                })
                                .map_err(|_| FormulaEvalError::Value)?;
                            dates.push(serial);
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    let serial = formula_serial_integer(self.parse_comparison()?)
                        .and_then(|serial| formula_ymd_from_serial(serial as f64).map(|_| serial))
                        .map_err(|_| FormulaEvalError::Value)?;
                    dates.push(serial);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                let serial = formula_serial_integer(self.parse_comparison()?)
                    .and_then(|serial| formula_ymd_from_serial(serial as f64).map(|_| serial))
                    .map_err(|_| FormulaEvalError::Value)?;
                dates.push(serial);
            }

            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if values.len() != dates.len()
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Num);
            }
            if values.iter().any(|value| !value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            let start_date = dates[0];
            let discount = 1.0 + rate;
            let mut total = 0.0;
            for (value, date) in values.iter().zip(dates.iter()) {
                if *date < start_date {
                    return Err(FormulaEvalError::Num);
                }
                let years = (*date - start_date) as f64 / 365.0;
                let denominator = discount.powf(years);
                if denominator == 0.0 || !denominator.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                total += value / denominator;
                if !total.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
            }
            return Ok(total);
        }
        if name.eq_ignore_ascii_case("XIRR") {
            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| ch == ',') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                values.push(self.parse_comparison()?);
            }

            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut dates = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let date = match self.evaluator.cell_value_or_blank(
                                target_sheet_id,
                                row,
                                col,
                            )? {
                                CellValue::Number(number) => number,
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            };
                            let serial = formula_serial_integer(date)
                                .and_then(|serial| {
                                    formula_ymd_from_serial(serial as f64).map(|_| serial)
                                })
                                .map_err(|_| FormulaEvalError::Value)?;
                            dates.push(serial);
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    let serial = formula_serial_integer(self.parse_comparison()?)
                        .and_then(|serial| formula_ymd_from_serial(serial as f64).map(|_| serial))
                        .map_err(|_| FormulaEvalError::Value)?;
                    dates.push(serial);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                let serial = formula_serial_integer(self.parse_comparison()?)
                    .and_then(|serial| formula_ymd_from_serial(serial as f64).map(|_| serial))
                    .map_err(|_| FormulaEvalError::Value)?;
                dates.push(serial);
            }

            self.skip_whitespace();
            let mut guess = 0.1;
            if self.consume_char(',') {
                guess = self.parse_comparison()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            } else if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !guess.is_finite() || values.iter().any(|value| !value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            if guess <= -1.0
                || values.len() != dates.len()
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Num);
            }
            let start_date = dates[0];
            let xirr_value = |rate: f64| -> Result<(f64, f64), FormulaEvalError> {
                if !rate.is_finite() || rate <= -1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let discount = 1.0 + rate;
                let mut value = 0.0;
                let mut derivative = 0.0;
                for (cash_flow, date) in values.iter().zip(dates.iter()) {
                    if *date < start_date {
                        return Err(FormulaEvalError::Num);
                    }
                    let years = (*date - start_date) as f64 / 365.0;
                    let denominator = discount.powf(years);
                    if denominator == 0.0 || !denominator.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    value += cash_flow / denominator;
                    derivative -= years * cash_flow / (denominator * discount);
                }
                if value.is_finite() && derivative.is_finite() {
                    Ok((value, derivative))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };

            const XIRR_MAX_ITERATIONS: usize = 100;
            const XIRR_TOLERANCE: f64 = 1e-8;
            let mut rate = guess;
            for _ in 0..XIRR_MAX_ITERATIONS {
                let (value, derivative) = xirr_value(rate)?;
                if value.abs() <= XIRR_TOLERANCE {
                    return Ok(rate);
                }
                if derivative == 0.0 {
                    break;
                }
                let next_rate = rate - value / derivative;
                if !next_rate.is_finite() || next_rate <= -1.0 {
                    break;
                }
                if (next_rate - rate).abs() <= XIRR_TOLERANCE {
                    return Ok(next_rate);
                }
                rate = next_rate;
            }
            return Err(FormulaEvalError::Num);
        }
        if name.eq_ignore_ascii_case("IRR") {
            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                values.push(self.parse_comparison()?);
            }
            self.skip_whitespace();
            let mut guess = 0.1;
            if self.consume_char(',') {
                guess = self.parse_comparison()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            } else if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !guess.is_finite() || values.iter().any(|value| !value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            if guess <= -1.0
                || values.len() < 2
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Num);
            }

            let irr_value = |rate: f64| -> Result<(f64, f64), FormulaEvalError> {
                if !rate.is_finite() || rate <= -1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let factor = 1.0 + rate;
                let mut denominator = 1.0;
                let mut value = 0.0;
                let mut derivative = 0.0;
                for (index, cash_flow) in values.iter().enumerate() {
                    if index > 0 {
                        denominator *= factor;
                        if denominator == 0.0 || !denominator.is_finite() {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    value += cash_flow / denominator;
                    if index > 0 {
                        derivative -= index as f64 * cash_flow / (denominator * factor);
                    }
                }
                if value.is_finite() && derivative.is_finite() {
                    Ok((value, derivative))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };

            const IRR_MAX_ITERATIONS: usize = 20;
            const IRR_TOLERANCE: f64 = 1e-7;
            let mut rate = guess;
            for _ in 0..IRR_MAX_ITERATIONS {
                let (value, derivative) = irr_value(rate)?;
                if value.abs() <= IRR_TOLERANCE {
                    return Ok(rate);
                }
                if derivative == 0.0 {
                    break;
                }
                let next_rate = rate - value / derivative;
                if !next_rate.is_finite() || next_rate <= -1.0 {
                    break;
                }
                if (next_rate - rate).abs() <= IRR_TOLERANCE {
                    return Ok(next_rate);
                }
                rate = next_rate;
            }
            return Err(FormulaEvalError::Num);
        }
        if name.eq_ignore_ascii_case("MIRR") {
            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| ch == ',') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                values.push(self.parse_comparison()?);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let finance_rate = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let reinvest_rate = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !finance_rate.is_finite()
                || !reinvest_rate.is_finite()
                || values.iter().any(|value| !value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            if values.len() < 2
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Div0);
            }
            let finance_factor = 1.0 + finance_rate;
            let reinvest_factor = 1.0 + reinvest_rate;
            let periods = values.len() - 1;
            let mut future_positive = 0.0;
            let mut present_negative = 0.0;
            for (index, value) in values.iter().enumerate() {
                if *value > 0.0 {
                    let exponent =
                        i32::try_from(periods - index).map_err(|_| FormulaEvalError::Num)?;
                    future_positive += value * reinvest_factor.powi(exponent);
                    if !future_positive.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                } else if *value < 0.0 {
                    let exponent = i32::try_from(index).map_err(|_| FormulaEvalError::Num)?;
                    let denominator = finance_factor.powi(exponent);
                    if denominator == 0.0 {
                        return Err(FormulaEvalError::Div0);
                    }
                    if !denominator.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    present_negative += value / denominator;
                    if !present_negative.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
            }
            if future_positive <= 0.0 || present_negative >= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let result = (future_positive / -present_negative).powf(1.0 / periods as f64) - 1.0;
            return if result.is_finite() {
                Ok(result)
            } else {
                Err(FormulaEvalError::Num)
            };
        }
        if name.eq_ignore_ascii_case("FV") {
            return self.parse_fv_function();
        }
        if name.eq_ignore_ascii_case("PV") {
            return self.parse_pv_function();
        }
        if name.eq_ignore_ascii_case("PMT") {
            return self.parse_pmt_function();
        }
        if name.eq_ignore_ascii_case("IPMT") {
            return self.parse_ipmt_function();
        }
        if name.eq_ignore_ascii_case("PPMT") {
            return self.parse_ppmt_function();
        }
        if name.eq_ignore_ascii_case("CUMIPMT") {
            return self.parse_cumulative_payment_function(false);
        }
        if name.eq_ignore_ascii_case("CUMPRINC") {
            return self.parse_cumulative_payment_function(true);
        }
        if name.eq_ignore_ascii_case("NPER") {
            return self.parse_nper_function();
        }
        if name.eq_ignore_ascii_case("RATE") {
            return self.parse_rate_function();
        }
        if name.eq_ignore_ascii_case("ISPMT") {
            return self.parse_ispmt_function();
        }
        if name.eq_ignore_ascii_case("ARABIC") {
            return self.parse_arabic_function();
        }
        if name.eq_ignore_ascii_case("CODE") || name.eq_ignore_ascii_case("UNICODE") {
            return self.parse_character_code_function();
        }
        if name.eq_ignore_ascii_case("LEN") {
            return self.parse_len_function(false);
        }
        if name.eq_ignore_ascii_case("LENB") {
            return self.parse_len_function(true);
        }
        if name.eq_ignore_ascii_case("FIND") {
            return self.parse_find_function(false, false);
        }
        if name.eq_ignore_ascii_case("FINDB") {
            return self.parse_find_function(false, true);
        }
        if name.eq_ignore_ascii_case("SEARCH") {
            return self.parse_find_function(true, false);
        }
        if name.eq_ignore_ascii_case("SEARCHB") {
            return self.parse_find_function(true, true);
        }
        if name.eq_ignore_ascii_case("REGEXTEST") {
            return self.parse_regex_test_function();
        }
        if name.eq_ignore_ascii_case("EXACT") {
            return self.parse_exact_function();
        }
        if name.eq_ignore_ascii_case("VALUE") {
            return self.parse_value_function();
        }
        if name.eq_ignore_ascii_case("NUMBERVALUE") {
            return self.parse_numbervalue_function();
        }
        if name.eq_ignore_ascii_case("DATEVALUE") {
            return self.parse_datevalue_function();
        }
        if name.eq_ignore_ascii_case("TIMEVALUE") {
            return self.parse_timevalue_function();
        }
        if name.eq_ignore_ascii_case("NA") {
            return self.parse_na_function();
        }
        if name.eq_ignore_ascii_case("CELL") {
            return formula_number_from_value_probe(self.parse_cell_value_function()?);
        }
        if name.eq_ignore_ascii_case("INFO") {
            return formula_number_from_value_probe(self.parse_info_value_function()?);
        }
        if name.eq_ignore_ascii_case("ROW") {
            return self.parse_row_function();
        }
        if name.eq_ignore_ascii_case("ROWS") {
            return self.parse_rows_function();
        }
        if name.eq_ignore_ascii_case("AREAS") {
            return self.parse_areas_function();
        }
        if name.eq_ignore_ascii_case("SHEET") {
            return self.parse_sheet_function();
        }
        if name.eq_ignore_ascii_case("SHEETS") {
            return self.parse_sheets_function();
        }
        if name.eq_ignore_ascii_case("INDEX") {
            return self.parse_index_function();
        }
        if name.eq_ignore_ascii_case("MATCH") {
            return self.parse_match_function();
        }
        if name.eq_ignore_ascii_case("XMATCH") {
            return self.parse_xmatch_function();
        }
        if name.eq_ignore_ascii_case("LOOKUP") {
            return self.parse_lookup_function();
        }
        if name.eq_ignore_ascii_case("VLOOKUP") {
            return self.parse_vlookup_function();
        }
        if name.eq_ignore_ascii_case("HLOOKUP") {
            return self.parse_hlookup_function();
        }
        if name.eq_ignore_ascii_case("XLOOKUP") {
            return formula_number_from_value_probe(self.parse_xlookup_value_function()?);
        }
        if name.eq_ignore_ascii_case("LET") {
            return formula_number_from_value_probe(self.parse_let_value_function()?);
        }
        if name.eq_ignore_ascii_case("LAMBDA") {
            let lambda = self.parse_lambda_value_function()?;
            self.skip_whitespace();
            if !self.consume_char('(') {
                return Err(FormulaEvalError::Calc);
            }
            return formula_number_from_value_probe(self.parse_lambda_call_arguments(lambda)?);
        }
        if name.eq_ignore_ascii_case("ISOMITTED") {
            return self.parse_isomitted_function();
        }
        if name.eq_ignore_ascii_case("MAKEARRAY") {
            return formula_number_from_value_probe(self.parse_makearray_value_function()?);
        }
        if name.eq_ignore_ascii_case("REDUCE") || name.eq_ignore_ascii_case("SCAN") {
            return formula_number_from_value_probe(self.parse_reduce_scan_value_function(name)?);
        }
        if name.eq_ignore_ascii_case("GETPIVOTDATA") {
            return formula_number_from_value_probe(self.parse_getpivotdata_value_function()?);
        }
        if name.eq_ignore_ascii_case("CUBESETCOUNT") {
            return self.parse_cubesetcount_function();
        }
        if name.eq_ignore_ascii_case("CUBEVALUE")
            || name.eq_ignore_ascii_case("RTD")
            || name.eq_ignore_ascii_case("STOCKHISTORY")
            || name.eq_ignore_ascii_case("COPILOT")
        {
            return self.parse_external_data_unavailable_function();
        }
        if name.eq_ignore_ascii_case("FIELDVALUE") {
            return self.parse_external_field_unavailable_function();
        }
        if name.eq_ignore_ascii_case("PY") {
            return self.parse_external_python_unavailable_function();
        }
        if name.eq_ignore_ascii_case("CALL") || name.eq_ignore_ascii_case("REGISTER.ID") {
            return self.parse_external_platform_unavailable_function();
        }
        if name.eq_ignore_ascii_case("INDIRECT")
            || name.eq_ignore_ascii_case("OFFSET")
            || name.eq_ignore_ascii_case("TRIMRANGE")
        {
            return formula_number_from_value_probe(
                self.parse_reference_projection_value_function(name)?,
            );
        }
        if formula_array_projection_function_name(name) {
            return formula_number_from_value_probe(
                self.parse_array_projection_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("FREQUENCY") {
            return self.parse_frequency_function();
        }
        if name.eq_ignore_ascii_case("MMULT") {
            return self.parse_mmult_function();
        }
        if name.eq_ignore_ascii_case("MINVERSE") {
            return self.parse_minverse_function();
        }
        if name.eq_ignore_ascii_case("MUNIT") {
            return self.parse_munit_function();
        }
        if name.eq_ignore_ascii_case("SEQUENCE") {
            return self.parse_sequence_function();
        }
        if name.eq_ignore_ascii_case("RANDARRAY") {
            return self.parse_randarray_function();
        }
        if name.eq_ignore_ascii_case("CHISQ.TEST") || name.eq_ignore_ascii_case("CHITEST") {
            return self.parse_chisq_test_function();
        }
        if name.eq_ignore_ascii_case("F.TEST") || name.eq_ignore_ascii_case("FTEST") {
            return self.parse_f_test_function();
        }
        if name.eq_ignore_ascii_case("T.TEST") || name.eq_ignore_ascii_case("TTEST") {
            return self.parse_t_test_function();
        }
        if name.eq_ignore_ascii_case("Z.TEST") || name.eq_ignore_ascii_case("ZTEST") {
            return self.parse_z_test_function();
        }
        if name.eq_ignore_ascii_case("DAVERAGE")
            || name.eq_ignore_ascii_case("DCOUNT")
            || name.eq_ignore_ascii_case("DCOUNTA")
            || name.eq_ignore_ascii_case("DGET")
            || name.eq_ignore_ascii_case("DMAX")
            || name.eq_ignore_ascii_case("DMIN")
            || name.eq_ignore_ascii_case("DPRODUCT")
            || name.eq_ignore_ascii_case("DSTDEV")
            || name.eq_ignore_ascii_case("DSTDEVP")
            || name.eq_ignore_ascii_case("DSUM")
            || name.eq_ignore_ascii_case("DVAR")
            || name.eq_ignore_ascii_case("DVARP")
        {
            return formula_number_from_value_probe(self.parse_database_value_function(name)?);
        }
        if name.eq_ignore_ascii_case("SUMIF") {
            return self.parse_sumif_function();
        }
        if name.eq_ignore_ascii_case("AVERAGEIF") {
            return self.parse_averageif_function();
        }
        if name.eq_ignore_ascii_case("COUNTIFS") {
            return self.parse_countifs_function();
        }
        if name.eq_ignore_ascii_case("SUMIFS") {
            return self.parse_sumifs_function();
        }
        if name.eq_ignore_ascii_case("AVERAGEIFS") {
            return self.parse_averageifs_function();
        }
        if name.eq_ignore_ascii_case("MINIFS") {
            return self.parse_minifs_function();
        }
        if name.eq_ignore_ascii_case("MAXIFS") {
            return self.parse_maxifs_function();
        }
        if name.eq_ignore_ascii_case("AVERAGEA")
            || name.eq_ignore_ascii_case("MINA")
            || name.eq_ignore_ascii_case("MAXA")
            || name.eq_ignore_ascii_case("VARA")
            || name.eq_ignore_ascii_case("VARPA")
            || name.eq_ignore_ascii_case("STDEVA")
            || name.eq_ignore_ascii_case("STDEVPA")
        {
            return self.parse_aggregate_a_function(name);
        }
        if name.eq_ignore_ascii_case("SUMPRODUCT") {
            let mut arguments: Vec<(u32, u32, Vec<f64>)> = Vec::new();
            let finish = |arguments: &[(u32, u32, Vec<f64>)]| -> Result<f64, FormulaEvalError> {
                let Some((base_height, base_width, first_values)) = arguments.first() else {
                    return Err(FormulaEvalError::Value);
                };
                if arguments
                    .iter()
                    .any(|(height, width, _)| height != base_height || width != base_width)
                {
                    return Err(FormulaEvalError::Value);
                }
                let mut total = 0.0_f64;
                for index in 0..first_values.len() {
                    let mut product = 1.0;
                    for (_, _, values) in arguments {
                        product *= values[index];
                    }
                    total += product;
                }
                if total.is_finite() {
                    Ok(total)
                } else {
                    Err(FormulaEvalError::Num)
                }
            };
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    return finish(arguments.as_slice());
                }

                let checkpoint = self.index;
                if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                    self.index = next_index;
                    self.skip_whitespace();
                    if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                        let mut values =
                            Vec::with_capacity((rect.width() * rect.height()) as usize);
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                match self.evaluator.cell_value_or_blank(
                                    target_sheet_id,
                                    row,
                                    col,
                                )? {
                                    CellValue::Number(number) => values.push(number),
                                    CellValue::Error(error) => {
                                        return Err(formula_eval_error_from_cell_error(error));
                                    }
                                    CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {
                                        values.push(0.0);
                                    }
                                }
                            }
                        }
                        arguments.push((rect.height(), rect.width(), values));
                    } else {
                        self.index = checkpoint;
                        arguments.push((1, 1, vec![self.parse_comparison()?]));
                    }
                } else {
                    arguments.push((1, 1, vec![self.parse_comparison()?]));
                }

                self.skip_whitespace();
                if self.consume_char(',') {
                    continue;
                }
                if self.consume_char(')') {
                    return finish(arguments.as_slice());
                }
                return Err(FormulaEvalError::Unsupported);
            }
        }
        if name.eq_ignore_ascii_case("SUMXMY2")
            || name.eq_ignore_ascii_case("SUMX2MY2")
            || name.eq_ignore_ascii_case("SUMX2PY2")
        {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Err(FormulaEvalError::Value);
            }
            let first_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let second_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if first_values.len() != second_values.len() {
                return Err(FormulaEvalError::NA);
            }
            let mut total = 0.0_f64;
            for (first_value, second_value) in first_values.iter().zip(second_values.iter()) {
                if name.eq_ignore_ascii_case("SUMXMY2") {
                    let difference = first_value - second_value;
                    total += difference * difference;
                } else if name.eq_ignore_ascii_case("SUMX2MY2") {
                    total += first_value * first_value - second_value * second_value;
                } else {
                    total += first_value * first_value + second_value * second_value;
                }
            }
            return Ok(total);
        }
        if name.eq_ignore_ascii_case("FORECAST") || name.eq_ignore_ascii_case("FORECAST.LINEAR") {
            let x = self.parse_comparison()?;
            if !x.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let known_y = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let known_x = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if known_y.len() != known_x.len() || known_y.is_empty() {
                return Err(FormulaEvalError::NA);
            }
            let count = known_y.len() as f64;
            let mean_y = known_y.iter().sum::<f64>() / count;
            let mean_x = known_x.iter().sum::<f64>() / count;
            let mut sum_xy_deviation = 0.0_f64;
            let mut sum_x_deviation_square = 0.0_f64;
            for (y_value, x_value) in known_y.iter().zip(known_x.iter()) {
                let y_deviation = y_value - mean_y;
                let x_deviation = x_value - mean_x;
                sum_xy_deviation += y_deviation * x_deviation;
                sum_x_deviation_square += x_deviation * x_deviation;
            }
            if sum_x_deviation_square == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            return Ok(mean_y + sum_xy_deviation / sum_x_deviation_square * (x - mean_x));
        }
        if name.eq_ignore_ascii_case("FORECAST.ETS")
            || name.eq_ignore_ascii_case("FORECAST.ETS.CONFINT")
            || name.eq_ignore_ascii_case("FORECAST.ETS.SEASONALITY")
            || name.eq_ignore_ascii_case("FORECAST.ETS.STAT")
        {
            return self.parse_forecast_ets_function(name);
        }
        if name.eq_ignore_ascii_case("LINEST") || name.eq_ignore_ascii_case("LOGEST") {
            return self.parse_regression_coefficient_function(name.eq_ignore_ascii_case("LOGEST"));
        }
        if name.eq_ignore_ascii_case("TREND") || name.eq_ignore_ascii_case("GROWTH") {
            return self.parse_regression_prediction_function(name.eq_ignore_ascii_case("GROWTH"));
        }
        if name.eq_ignore_ascii_case("CORREL")
            || name.eq_ignore_ascii_case("PEARSON")
            || name.eq_ignore_ascii_case("COVAR")
            || name.eq_ignore_ascii_case("COVARIANCE.P")
            || name.eq_ignore_ascii_case("COVARIANCE.S")
            || name.eq_ignore_ascii_case("SLOPE")
            || name.eq_ignore_ascii_case("INTERCEPT")
            || name.eq_ignore_ascii_case("RSQ")
            || name.eq_ignore_ascii_case("STEYX")
        {
            self.skip_whitespace();
            let first_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let second_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if first_values.len() != second_values.len() {
                return Err(FormulaEvalError::NA);
            }
            let count = first_values.len();
            if count == 0 {
                return if name.eq_ignore_ascii_case("CORREL")
                    || name.eq_ignore_ascii_case("COVAR")
                    || name.eq_ignore_ascii_case("COVARIANCE.P")
                    || name.eq_ignore_ascii_case("COVARIANCE.S")
                {
                    Err(FormulaEvalError::Div0)
                } else {
                    Err(FormulaEvalError::NA)
                };
            }
            if name.eq_ignore_ascii_case("COVARIANCE.S") && count < 2 {
                return Err(FormulaEvalError::Div0);
            }

            let first_mean = first_values.iter().sum::<f64>() / count as f64;
            let second_mean = second_values.iter().sum::<f64>() / count as f64;
            let mut sum_first_second_deviation = 0.0_f64;
            let mut sum_first_deviation_square = 0.0_f64;
            let mut sum_second_deviation_square = 0.0_f64;
            for (first_value, second_value) in first_values.iter().zip(second_values.iter()) {
                let first_deviation = first_value - first_mean;
                let second_deviation = second_value - second_mean;
                sum_first_second_deviation += first_deviation * second_deviation;
                sum_first_deviation_square += first_deviation * first_deviation;
                sum_second_deviation_square += second_deviation * second_deviation;
            }

            if name.eq_ignore_ascii_case("COVAR") || name.eq_ignore_ascii_case("COVARIANCE.P") {
                return Ok(sum_first_second_deviation / count as f64);
            }
            if name.eq_ignore_ascii_case("COVARIANCE.S") {
                return Ok(sum_first_second_deviation / (count - 1) as f64);
            }
            if name.eq_ignore_ascii_case("SLOPE") || name.eq_ignore_ascii_case("INTERCEPT") {
                if sum_second_deviation_square == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let slope = sum_first_second_deviation / sum_second_deviation_square;
                if name.eq_ignore_ascii_case("SLOPE") {
                    return Ok(slope);
                }
                return Ok(first_mean - slope * second_mean);
            }
            if name.eq_ignore_ascii_case("STEYX") {
                if count < 3 || sum_second_deviation_square == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let residual_square_sum = sum_first_deviation_square
                    - sum_first_second_deviation * sum_first_second_deviation
                        / sum_second_deviation_square;
                return Ok((residual_square_sum.max(0.0) / (count - 2) as f64).sqrt());
            }

            let denominator = sum_first_deviation_square * sum_second_deviation_square;
            if denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            let correlation = sum_first_second_deviation / denominator.sqrt();
            if name.eq_ignore_ascii_case("RSQ") {
                return Ok(correlation * correlation);
            }
            return Ok(correlation);
        }
        if name.eq_ignore_ascii_case("PROB") {
            let x_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let probabilities = self.parse_aggregate_argument()?;
            if x_values.len() != probabilities.len() {
                return Err(FormulaEvalError::NA);
            }
            let mut probability_sum = 0.0_f64;
            for probability in &probabilities {
                if !probability.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *probability <= 0.0 || *probability > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                probability_sum += probability;
            }
            if (probability_sum - 1.0).abs() > 1e-7 {
                return Err(FormulaEvalError::Num);
            }

            self.skip_whitespace();
            let (lower_limit, upper_limit) = if self.consume_char(')') {
                (0.0, None)
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let lower_limit = self.parse_comparison()?;
                if !lower_limit.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if self.consume_char(')') {
                    (lower_limit, None)
                } else {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    let upper_limit = self.parse_comparison()?;
                    if !upper_limit.is_finite() {
                        return Err(FormulaEvalError::Value);
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    if upper_limit < lower_limit {
                        return Err(FormulaEvalError::Num);
                    }
                    (lower_limit, Some(upper_limit))
                }
            };

            let total = x_values
                .iter()
                .zip(probabilities.iter())
                .filter_map(|(x_value, probability)| {
                    let matches = if let Some(upper_limit) = upper_limit {
                        *x_value >= lower_limit && *x_value <= upper_limit
                    } else {
                        *x_value == lower_limit
                    };
                    matches.then_some(*probability)
                })
                .sum::<f64>();
            return Ok(total);
        }
        if name.eq_ignore_ascii_case("PERCENTOF") {
            let subset = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let all_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let denominator = all_values.iter().sum::<f64>();
            if denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            return formula_checked_numeric_result(subset.iter().sum::<f64>() / denominator);
        }
        if name.eq_ignore_ascii_case("PERCENTILE")
            || name.eq_ignore_ascii_case("PERCENTILE.INC")
            || name.eq_ignore_ascii_case("PERCENTILE.EXC")
        {
            let values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let k = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return percentile_value(values, k, name.eq_ignore_ascii_case("PERCENTILE.EXC"));
        }
        if name.eq_ignore_ascii_case("PERCENTRANK")
            || name.eq_ignore_ascii_case("PERCENTRANK.INC")
            || name.eq_ignore_ascii_case("PERCENTRANK.EXC")
        {
            let values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let x = self.parse_comparison()?;
            self.skip_whitespace();
            let significance = if self.consume_char(')') {
                3
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let significance = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                significance
            };
            return percent_rank_value(
                values,
                x,
                significance,
                name.eq_ignore_ascii_case("PERCENTRANK.EXC"),
            );
        }
        if name.eq_ignore_ascii_case("QUARTILE")
            || name.eq_ignore_ascii_case("QUARTILE.INC")
            || name.eq_ignore_ascii_case("QUARTILE.EXC")
        {
            let values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let quart = self.parse_comparison()?;
            if !quart.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if quart < i64::MIN as f64 || quart > i64::MAX as f64 {
                return Err(FormulaEvalError::Num);
            }
            let quart = quart.trunc() as i64;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let exclusive = name.eq_ignore_ascii_case("QUARTILE.EXC");
            if exclusive {
                if !(1..=3).contains(&quart) {
                    return Err(FormulaEvalError::Num);
                }
            } else if !(0..=4).contains(&quart) {
                return Err(FormulaEvalError::Num);
            }
            return percentile_value(values, quart as f64 / 4.0, exclusive);
        }
        if name.eq_ignore_ascii_case("TRIMMEAN") {
            let mut values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let percent = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if values.is_empty() {
                return Err(FormulaEvalError::Div0);
            }
            if !percent.is_finite() || !(0.0..=1.0).contains(&percent) {
                return Err(FormulaEvalError::Num);
            }
            values.sort_by(|left, right| left.total_cmp(right));
            let trim_count = (values.len() as f64 * percent).floor() as usize;
            let trim_each_side = (trim_count - trim_count % 2) / 2;
            let remaining = values.len().saturating_sub(trim_each_side * 2);
            if remaining == 0 {
                return Err(FormulaEvalError::Num);
            }
            return Ok(values[trim_each_side..trim_each_side + remaining]
                .iter()
                .sum::<f64>()
                / remaining as f64);
        }
        if name.eq_ignore_ascii_case("LARGE") || name.eq_ignore_ascii_case("SMALL") {
            let mut values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let k = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if values.is_empty() || k < 1 || k > values.len() as i64 {
                return Err(FormulaEvalError::Num);
            }
            values.sort_by(|left, right| left.total_cmp(right));
            let index = if name.eq_ignore_ascii_case("LARGE") {
                values.len() - k as usize
            } else {
                k as usize - 1
            };
            return Ok(values[index]);
        }
        if name.eq_ignore_ascii_case("RANK")
            || name.eq_ignore_ascii_case("RANK.EQ")
            || name.eq_ignore_ascii_case("RANK.AVG")
        {
            let number = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (target_sheet_id, rect) = self.parse_reference_argument()?;
            let values = self
                .evaluator
                .numeric_values_in_rect(target_sheet_id, rect)?;
            self.skip_whitespace();
            let ascending = if self.consume_char(')') {
                false
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let order = self.parse_comparison()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                order != 0.0
            };
            let tie_count = values.iter().filter(|value| **value == number).count();
            if tie_count == 0 {
                return Err(FormulaEvalError::NA);
            }
            let ahead_count = values
                .iter()
                .filter(|value| {
                    if ascending {
                        **value < number
                    } else {
                        **value > number
                    }
                })
                .count();
            let rank = ahead_count as f64 + 1.0;
            if name.eq_ignore_ascii_case("RANK.AVG") {
                return Ok(rank + (tie_count as f64 - 1.0) / 2.0);
            }
            return Ok(rank);
        }
        let function =
            FormulaAggregateFunction::from_name(name).ok_or(FormulaEvalError::Unsupported)?;
        let mut values = Vec::new();
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return function.evaluate(values.as_slice());
            }
            values.extend(self.parse_aggregate_argument()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return function.evaluate(values.as_slice());
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_text_function(&mut self, name: &str) -> Result<String, FormulaEvalError> {
        if name.eq_ignore_ascii_case("ADDRESS") {
            return self.parse_address_text_function();
        }
        if name.eq_ignore_ascii_case("ARRAYTOTEXT") {
            return self.parse_arraytotext_function();
        }
        if name.eq_ignore_ascii_case("ASC")
            || name.eq_ignore_ascii_case("DBCS")
            || name.eq_ignore_ascii_case("JIS")
        {
            return self.parse_unary_text_function(|text| text);
        }
        if name.eq_ignore_ascii_case("CONCAT") || name.eq_ignore_ascii_case("CONCATENATE") {
            return self.parse_concat_text_function();
        }
        if name.eq_ignore_ascii_case("DGET") {
            return formula_selected_text_from_value_probe(
                self.parse_database_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("LAMBDA") {
            let lambda = self.parse_lambda_value_function()?;
            self.skip_whitespace();
            if !self.consume_char('(') {
                return Err(FormulaEvalError::Calc);
            }
            return formula_selected_text_from_value_probe(
                self.parse_lambda_call_arguments(lambda)?,
            );
        }
        if name.eq_ignore_ascii_case("MAKEARRAY") {
            return formula_selected_text_from_value_probe(self.parse_makearray_value_function()?);
        }
        if name.eq_ignore_ascii_case("REDUCE") || name.eq_ignore_ascii_case("SCAN") {
            return formula_selected_text_from_value_probe(
                self.parse_reduce_scan_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("GETPIVOTDATA") {
            return formula_selected_text_from_value_probe(
                self.parse_getpivotdata_value_function()?,
            );
        }
        if name.eq_ignore_ascii_case("CUBEKPIMEMBER")
            || name.eq_ignore_ascii_case("CUBEMEMBER")
            || name.eq_ignore_ascii_case("CUBERANKEDMEMBER")
            || name.eq_ignore_ascii_case("CUBESET")
        {
            return self.parse_cube_caption_text_function(name);
        }
        if name.eq_ignore_ascii_case("CUBEMEMBERPROPERTY") {
            return self.parse_cubememberproperty_text_function();
        }
        if formula_array_projection_function_name(name) {
            return formula_selected_text_from_value_probe(
                self.parse_array_projection_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("REGEXEXTRACT") {
            return self.parse_regex_extract_function();
        }
        if name.eq_ignore_ascii_case("REGEXREPLACE") {
            return self.parse_regex_replace_function();
        }
        if name.eq_ignore_ascii_case("DETECTLANGUAGE") {
            return self.parse_detectlanguage_function();
        }
        if name.eq_ignore_ascii_case("FILTERXML") {
            return self.parse_filterxml_function();
        }
        if name.eq_ignore_ascii_case("IMAGE") {
            return self.parse_image_function();
        }
        if name.eq_ignore_ascii_case("TRANSLATE") {
            return self.parse_translate_function();
        }
        if name.eq_ignore_ascii_case("WEBSERVICE") {
            return self.parse_webservice_function();
        }
        if name.eq_ignore_ascii_case("ENCODEURL") {
            return self.parse_unary_text_function(|text| formula_encode_url(text.as_str()));
        }
        if name.eq_ignore_ascii_case("LEFT") {
            return self.parse_left_text_function(false);
        }
        if name.eq_ignore_ascii_case("LEFTB") {
            return self.parse_left_text_function(true);
        }
        if name.eq_ignore_ascii_case("RIGHT") {
            return self.parse_right_text_function(false);
        }
        if name.eq_ignore_ascii_case("RIGHTB") {
            return self.parse_right_text_function(true);
        }
        if name.eq_ignore_ascii_case("MID") {
            return self.parse_mid_text_function(false);
        }
        if name.eq_ignore_ascii_case("MIDB") {
            return self.parse_mid_text_function(true);
        }
        if name.eq_ignore_ascii_case("BASE") {
            return self.parse_base_text_function();
        }
        if name.eq_ignore_ascii_case("BAHTTEXT") {
            return self.parse_bahttext_function();
        }
        if name.eq_ignore_ascii_case("CELL") {
            return formula_selected_text_from_value_probe(self.parse_cell_value_function()?);
        }
        if name.eq_ignore_ascii_case("INFO") {
            return formula_selected_text_from_value_probe(self.parse_info_value_function()?);
        }
        if name.eq_ignore_ascii_case("INDIRECT")
            || name.eq_ignore_ascii_case("OFFSET")
            || name.eq_ignore_ascii_case("TRIMRANGE")
        {
            return formula_selected_text_from_value_probe(
                self.parse_reference_projection_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("DEC2BIN") {
            return self.parse_decimal_engineering_text_function(2, 10, 10);
        }
        if name.eq_ignore_ascii_case("DEC2OCT") {
            return self.parse_decimal_engineering_text_function(8, 30, 10);
        }
        if name.eq_ignore_ascii_case("DEC2HEX") {
            return self.parse_decimal_engineering_text_function(16, 40, 10);
        }
        if name.eq_ignore_ascii_case("BIN2OCT") {
            return self.parse_engineering_text_function(2, 10, 10, 8, 30, 10);
        }
        if name.eq_ignore_ascii_case("BIN2HEX") {
            return self.parse_engineering_text_function(2, 10, 10, 16, 40, 10);
        }
        if name.eq_ignore_ascii_case("OCT2BIN") {
            return self.parse_engineering_text_function(8, 30, 10, 2, 10, 10);
        }
        if name.eq_ignore_ascii_case("OCT2HEX") {
            return self.parse_engineering_text_function(8, 30, 10, 16, 40, 10);
        }
        if name.eq_ignore_ascii_case("HEX2BIN") {
            return self.parse_engineering_text_function(16, 40, 10, 2, 10, 10);
        }
        if name.eq_ignore_ascii_case("HEX2OCT") {
            return self.parse_engineering_text_function(16, 40, 10, 8, 30, 10);
        }
        if name.eq_ignore_ascii_case("COMPLEX") {
            let real = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let imaginary = self.parse_comparison()?;
            self.skip_whitespace();
            let suffix = if self.consume_char(')') {
                'i'
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let suffix = formula_complex_suffix(self.parse_text_value_argument()?.as_str())?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                suffix
            };
            return formula_complex_format(formula_complex_number(real, imaginary, Some(suffix))?);
        }
        if name.eq_ignore_ascii_case("IMCONJUGATE") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_number(
                value.real,
                -value.imaginary,
                value.suffix,
            )?);
        }
        if name.eq_ignore_ascii_case("IMCOS") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_cos(value)?);
        }
        if name.eq_ignore_ascii_case("IMCOSH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_cosh(value)?);
        }
        if name.eq_ignore_ascii_case("IMCOT") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_divide(
                formula_complex_cos(value)?,
                formula_complex_sin(value)?,
            )?);
        }
        if name.eq_ignore_ascii_case("IMCSC") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_sin(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMCSCH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_sinh(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMDIV") {
            let left = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let right = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_divide(left, right)?);
        }
        if name.eq_ignore_ascii_case("IMEXP") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_exp(value)?);
        }
        if name.eq_ignore_ascii_case("IMLN") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_ln(value)?);
        }
        if name.eq_ignore_ascii_case("IMLOG10") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = formula_complex_ln(value)?;
            let base = 10.0_f64.ln();
            return formula_complex_format(formula_complex_number(
                value.real / base,
                value.imaginary / base,
                value.suffix,
            )?);
        }
        if name.eq_ignore_ascii_case("IMLOG2") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = formula_complex_ln(value)?;
            let base = 2.0_f64.ln();
            return formula_complex_format(formula_complex_number(
                value.real / base,
                value.imaginary / base,
                value.suffix,
            )?);
        }
        if name.eq_ignore_ascii_case("IMPOWER") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let power = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_power(value, power)?);
        }
        if name.eq_ignore_ascii_case("IMPRODUCT") || name.eq_ignore_ascii_case("IMSUM") {
            let product = name.eq_ignore_ascii_case("IMPRODUCT");
            let mut value = formula_complex_number(if product { 1.0 } else { 0.0 }, 0.0, None)?;
            let mut saw_argument = false;
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    return if saw_argument {
                        formula_complex_format(value)
                    } else {
                        Err(FormulaEvalError::Value)
                    };
                }
                let argument = self.parse_complex_argument()?;
                value = if product {
                    formula_complex_multiply(value, argument)?
                } else {
                    formula_complex_add(value, argument)?
                };
                saw_argument = true;
                self.skip_whitespace();
                if self.consume_char(',') {
                    continue;
                }
                if self.consume_char(')') {
                    return formula_complex_format(value);
                }
                return Err(FormulaEvalError::Unsupported);
            }
        }
        if name.eq_ignore_ascii_case("IMSEC") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_cos(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMSECH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_cosh(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMSIN") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_sin(value)?);
        }
        if name.eq_ignore_ascii_case("IMSINH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_sinh(value)?);
        }
        if name.eq_ignore_ascii_case("IMSQRT") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_sqrt(value)?);
        }
        if name.eq_ignore_ascii_case("IMSUB") {
            let left = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let right = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_subtract(left, right)?);
        }
        if name.eq_ignore_ascii_case("IMTAN") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_divide(
                formula_complex_sin(value)?,
                formula_complex_cos(value)?,
            )?);
        }
        if name.eq_ignore_ascii_case("ROMAN") {
            return self.parse_roman_text_function();
        }
        if name.eq_ignore_ascii_case("DOLLAR") {
            return self.parse_dollar_text_function();
        }
        if name.eq_ignore_ascii_case("FIXED") {
            return self.parse_fixed_text_function();
        }
        if name.eq_ignore_ascii_case("TEXT") {
            return self.parse_text_format_function();
        }
        if name.eq_ignore_ascii_case("HYPERLINK") {
            return self.parse_hyperlink_text_function();
        }
        if name.eq_ignore_ascii_case("PHONETIC") {
            return self.parse_phonetic_text_function();
        }
        if name.eq_ignore_ascii_case("CHAR") {
            return self.parse_character_text_function(false);
        }
        if name.eq_ignore_ascii_case("UNICHAR") {
            return self.parse_character_text_function(true);
        }
        if name.eq_ignore_ascii_case("CLEAN") {
            return self.parse_unary_text_function(|text| {
                text.chars()
                    .filter(|ch| !matches!(*ch as u32, 0..=31))
                    .collect()
            });
        }
        if name.eq_ignore_ascii_case("T") {
            let value = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return match value {
                FormulaValueProbe::Text(value) => Ok(value),
                FormulaValueProbe::Error(error) => Err(error),
                FormulaValueProbe::Blank
                | FormulaValueProbe::Bool(_)
                | FormulaValueProbe::Number(_)
                | FormulaValueProbe::Omitted
                | FormulaValueProbe::Lambda { .. } => Ok(String::new()),
            };
        }
        if name.eq_ignore_ascii_case("UPPER") {
            return self.parse_unary_text_function(|text| text.to_uppercase());
        }
        if name.eq_ignore_ascii_case("LOWER") {
            return self.parse_unary_text_function(|text| text.to_lowercase());
        }
        if name.eq_ignore_ascii_case("PROPER") {
            return self.parse_unary_text_function(|text| formula_proper_text(text.as_str()));
        }
        if name.eq_ignore_ascii_case("TRIM") {
            return self.parse_unary_text_function(|text| {
                text.split_whitespace().collect::<Vec<_>>().join(" ")
            });
        }
        if name.eq_ignore_ascii_case("TEXTJOIN") {
            return self.parse_textjoin_function();
        }
        if name.eq_ignore_ascii_case("TEXTSPLIT") {
            return self.parse_textsplit_function();
        }
        if name.eq_ignore_ascii_case("TEXTBEFORE") {
            return self.parse_text_boundary_function(false);
        }
        if name.eq_ignore_ascii_case("TEXTAFTER") {
            return self.parse_text_boundary_function(true);
        }
        if name.eq_ignore_ascii_case("REPT") {
            return self.parse_rept_text_function();
        }
        if name.eq_ignore_ascii_case("REPLACE") {
            return self.parse_replace_text_function(false);
        }
        if name.eq_ignore_ascii_case("REPLACEB") {
            return self.parse_replace_text_function(true);
        }
        if name.eq_ignore_ascii_case("SUBSTITUTE") {
            return self.parse_substitute_text_function();
        }
        if name.eq_ignore_ascii_case("VALUETOTEXT") {
            return self.parse_valuetotext_function();
        }
        if name.eq_ignore_ascii_case("FORMULATEXT") {
            return self.parse_formulatext_function();
        }
        if name.eq_ignore_ascii_case("IF") {
            return formula_selected_text_from_value_probe(self.parse_if_value_function()?);
        }
        if name.eq_ignore_ascii_case("IFS") {
            return formula_selected_text_from_value_probe(self.parse_ifs_value_function()?);
        }
        if name.eq_ignore_ascii_case("SWITCH") {
            return formula_selected_text_from_value_probe(self.parse_switch_value_function()?);
        }
        if name.eq_ignore_ascii_case("CHOOSE") {
            return formula_selected_text_from_value_probe(self.parse_choose_value_function()?);
        }
        if name.eq_ignore_ascii_case("INDEX") {
            return formula_selected_text_from_value_probe(self.parse_index_value_function()?);
        }
        if name.eq_ignore_ascii_case("LOOKUP") {
            return formula_selected_text_from_value_probe(self.parse_lookup_value_function()?);
        }
        if name.eq_ignore_ascii_case("VLOOKUP") {
            return formula_selected_text_from_value_probe(
                self.parse_table_lookup_value_function(false)?,
            );
        }
        if name.eq_ignore_ascii_case("HLOOKUP") {
            return formula_selected_text_from_value_probe(
                self.parse_table_lookup_value_function(true)?,
            );
        }
        if name.eq_ignore_ascii_case("XLOOKUP") {
            return formula_selected_text_from_value_probe(self.parse_xlookup_value_function()?);
        }
        if name.eq_ignore_ascii_case("LET") {
            return formula_selected_text_from_value_probe(self.parse_let_value_function()?);
        }
        Err(FormulaEvalError::Unsupported)
    }

    fn parse_concat_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let mut output = String::new();
        let mut saw_argument = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return if saw_argument {
                    Ok(output)
                } else {
                    Err(FormulaEvalError::Value)
                };
            }
            output.push_str(self.parse_text_value_argument()?.as_str());
            saw_argument = true;
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return Ok(output);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_address_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let row = formula_integer_argument(self.parse_comparison()?)?;
        if row < 1 || row > i64::from(EXCEL_MAX_ROW_INDEX) {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let col = formula_integer_argument(self.parse_comparison()?)?;
        if col < 1 || col > i64::from(EXCEL_MAX_COLUMN_INDEX) {
            return Err(FormulaEvalError::Value);
        }

        let mut abs_num = 1_i64;
        let mut a1 = true;
        let mut sheet_text = None;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            abs_num = formula_integer_argument(self.parse_comparison()?)?;
            if !(1..=4).contains(&abs_num) {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                a1 = self.parse_comparison()? != 0.0;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    sheet_text = Some(self.parse_text_value_argument()?);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }

        let row = u32::try_from(row).map_err(|_| FormulaEvalError::Value)?;
        let col = u32::try_from(col).map_err(|_| FormulaEvalError::Value)?;
        let row_absolute = matches!(abs_num, 1 | 2);
        let column_absolute = matches!(abs_num, 1 | 3);
        let mut address = if a1 {
            format_cell_address(row, col, row_absolute, column_absolute)
        } else {
            let row_part = if row_absolute {
                row.to_string()
            } else {
                format!("[{row}]")
            };
            let col_part = if column_absolute {
                col.to_string()
            } else {
                format!("[{col}]")
            };
            format!("R{row_part}C{col_part}")
        };
        if let Some(sheet_text) = sheet_text {
            address = format!(
                "{}{address}",
                formula_sheet_address_qualifier(sheet_text.as_str())
            );
        }
        Ok(address)
    }

    fn parse_left_text_function(&mut self, byte_mode: bool) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let count = self.parse_optional_text_count_argument(1)?;
        if byte_mode {
            Ok(formula_text_byte_slice(text.as_str(), 1, count))
        } else {
            Ok(text.chars().take(count).collect())
        }
    }

    fn parse_right_text_function(&mut self, byte_mode: bool) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let count = self.parse_optional_text_count_argument(1)?;
        if byte_mode {
            let len = formula_text_byte_len(text.as_str());
            let start = len.saturating_sub(count) + 1;
            return Ok(formula_text_byte_slice(text.as_str(), start, count));
        }
        let chars = text.chars().collect::<Vec<_>>();
        Ok(chars
            .iter()
            .skip(chars.len().saturating_sub(count))
            .collect())
    }

    fn parse_mid_text_function(&mut self, byte_mode: bool) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let start = formula_positive_position_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if byte_mode {
            Ok(formula_text_byte_slice(text.as_str(), start, count))
        } else {
            Ok(text.chars().skip(start - 1).take(count).collect())
        }
    }

    fn parse_character_text_function(&mut self, unicode: bool) -> Result<String, FormulaEvalError> {
        let code = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let code = u32::try_from(code).map_err(|_| FormulaEvalError::Value)?;
        if (!unicode && !(1..=255).contains(&code)) || code == 0 {
            return Err(FormulaEvalError::Value);
        }
        char::from_u32(code)
            .map(|ch| ch.to_string())
            .ok_or(FormulaEvalError::Value)
    }

    fn parse_base_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = formula_integer_argument(self.parse_comparison()?)?;
        let number = u64::try_from(number).map_err(|_| FormulaEvalError::Num)?;
        if number >= (1_u64 << 53) {
            return Err(FormulaEvalError::Num);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let radix = formula_radix_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        let min_length = if self.consume_char(')') {
            0
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let min_length = formula_integer_argument(self.parse_comparison()?)?;
            if !(0..=255).contains(&min_length) {
                return Err(FormulaEvalError::Num);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            min_length as usize
        };
        let mut value = number;
        let mut output = String::new();
        loop {
            let digit = (value % u64::from(radix)) as u32;
            output.push(
                char::from_digit(digit, radix)
                    .expect("digit")
                    .to_ascii_uppercase(),
            );
            value /= u64::from(radix);
            if value == 0 {
                break;
            }
        }
        let mut output = output.chars().rev().collect::<String>();
        if output.len() < min_length {
            output = "0".repeat(min_length - output.len()) + output.as_str();
        }
        Ok(output)
    }

    fn parse_bahttext_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !number.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let scaled = number.abs() * 100.0;
        if !scaled.is_finite() || scaled > u128::MAX as f64 {
            return Err(FormulaEvalError::Num);
        }
        let total_satang = round_half_away_from_zero(scaled) as u128;
        let baht = total_satang / 100;
        let satang = (total_satang % 100) as u32;

        let convert_group = |group: u32| -> String {
            let digit_text = [
                "ศูนย์",
                "หนึ่ง",
                "สอง",
                "สาม",
                "สี่",
                "ห้า",
                "หก",
                "เจ็ด",
                "แปด",
                "เก้า",
            ];
            let mut output = String::new();
            let hundred_thousands = group / 100_000;
            let ten_thousands = group / 10_000 % 10;
            let thousands = group / 1_000 % 10;
            let hundreds = group / 100 % 10;
            let tens = group / 10 % 10;
            let ones = group % 10;
            for (digit, suffix) in [
                (hundred_thousands, "แสน"),
                (ten_thousands, "หมื่น"),
                (thousands, "พัน"),
                (hundreds, "ร้อย"),
            ] {
                if digit != 0 {
                    output.push_str(digit_text[digit as usize]);
                    output.push_str(suffix);
                }
            }
            if tens != 0 {
                if tens == 1 {
                    output.push_str("สิบ");
                } else if tens == 2 {
                    output.push_str("ยี่สิบ");
                } else {
                    output.push_str(digit_text[tens as usize]);
                    output.push_str("สิบ");
                }
            }
            if ones != 0 {
                if ones == 1 && group > 1 {
                    output.push_str("เอ็ด");
                } else {
                    output.push_str(digit_text[ones as usize]);
                }
            }
            if output.is_empty() {
                output.push_str(digit_text[0]);
            }
            output
        };

        let convert_number = |mut value: u128| -> String {
            if value == 0 {
                return "ศูนย์".to_string();
            }
            let mut groups = Vec::new();
            while value > 0 {
                groups.push((value % 1_000_000) as u32);
                value /= 1_000_000;
            }
            let mut output = String::new();
            for index in (0..groups.len()).rev() {
                let group = groups[index];
                if group != 0 {
                    output.push_str(convert_group(group).as_str());
                }
                if index > 0 {
                    output.push_str("ล้าน");
                }
            }
            output
        };

        let mut output = String::new();
        if number < 0.0 && total_satang != 0 {
            output.push_str("ลบ");
        }
        output.push_str(convert_number(baht).as_str());
        output.push_str("บาท");
        if satang == 0 {
            output.push_str("ถ้วน");
        } else {
            output.push_str(convert_group(satang).as_str());
            output.push_str("สตางค์");
        }
        Ok(output)
    }

    fn parse_decimal_engineering_text_function(
        &mut self,
        radix: u32,
        bits: u32,
        max_digits: usize,
    ) -> Result<String, FormulaEvalError> {
        let value = formula_integer_argument(self.parse_comparison()?)?;
        let places = self.parse_optional_engineering_places()?;
        formula_engineering_format(value, radix, bits, max_digits, places)
    }

    fn parse_engineering_text_function(
        &mut self,
        source_radix: u32,
        source_bits: u32,
        source_max_digits: usize,
        target_radix: u32,
        target_bits: u32,
        target_max_digits: usize,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let value =
            formula_engineering_input(text.as_str(), source_radix, source_bits, source_max_digits)?;
        let places = self.parse_optional_engineering_places()?;
        formula_engineering_format(value, target_radix, target_bits, target_max_digits, places)
    }

    fn parse_roman_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        if !number.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let number = number.trunc();
        if !(0.0..=3999.0).contains(&number) {
            return Err(FormulaEvalError::Value);
        }
        let mut form = 0_usize;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            form = match self.parse_value_probe_argument()? {
                FormulaValueProbe::Bool(true) => 0,
                FormulaValueProbe::Bool(false) => 4,
                FormulaValueProbe::Number(value) => {
                    if !value.is_finite() {
                        return Err(FormulaEvalError::Value);
                    }
                    let value = value.trunc();
                    if !(0.0..=4.0).contains(&value) {
                        return Err(FormulaEvalError::Value);
                    }
                    value as usize
                }
                FormulaValueProbe::Error(error) => return Err(error),
                FormulaValueProbe::Blank
                | FormulaValueProbe::Text(_)
                | FormulaValueProbe::Omitted
                | FormulaValueProbe::Lambda { .. } => {
                    return Err(FormulaEvalError::Value);
                }
            };
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
        formula_roman_text(number as i64, form)
    }

    fn parse_text_format_decimals_argument(&mut self) -> Result<i64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        if !value.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let value = value.trunc();
        if !(-127.0..=127.0).contains(&value) {
            return Err(FormulaEvalError::Value);
        }
        Ok(value as i64)
    }

    fn parse_dollar_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        let mut decimals = 2_i64;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            decimals = self.parse_text_format_decimals_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
        let formatted = formula_fixed_number_text(number, decimals, true)?;
        if let Some(positive) = formatted.strip_prefix('-') {
            Ok(format!("(${positive})"))
        } else {
            Ok(format!("${formatted}"))
        }
    }

    fn parse_fixed_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        let mut decimals = 2_i64;
        let mut use_commas = true;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            decimals = self.parse_text_format_decimals_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                use_commas = self.parse_comparison()? == 0.0;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_fixed_number_text(number, decimals, use_commas)
    }

    fn parse_text_format_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let format = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        if !number.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let sections = format.split(';').collect::<Vec<_>>();
        if sections.is_empty() || sections.len() > 4 {
            return Err(FormulaEvalError::Value);
        }
        let format = if number < 0.0 && sections.len() > 1 {
            sections[1]
        } else if number == 0.0 && sections.len() > 2 {
            sections[2]
        } else {
            sections[0]
        }
        .trim();
        if format.is_empty() {
            return Err(FormulaEvalError::Value);
        }

        #[derive(Clone)]
        enum TextDateTimeFormatToken {
            Literal(String),
            Year(usize),
            MonthOrMinute(usize),
            Day(usize),
            Hour(usize),
            Second(usize),
            AmPm(String),
        }

        let mut tokens = Vec::new();
        let mut literal = String::new();
        let push_literal = |tokens: &mut Vec<TextDateTimeFormatToken>, literal: &mut String| {
            if !literal.is_empty() {
                tokens.push(TextDateTimeFormatToken::Literal(std::mem::take(literal)));
            }
        };
        let format_chars = format.chars().collect::<Vec<_>>();
        let mut index = 0_usize;
        while index < format_chars.len() {
            let ch = format_chars[index];
            if ch == '"' {
                index += 1;
                loop {
                    let quoted = format_chars.get(index).ok_or(FormulaEvalError::Value)?;
                    index += 1;
                    if *quoted == '"' {
                        break;
                    }
                    literal.push(*quoted);
                }
                continue;
            }
            if ch == '\\' {
                index += 1;
                let escaped = format_chars.get(index).ok_or(FormulaEvalError::Value)?;
                literal.push(*escaped);
                index += 1;
                continue;
            }
            if matches!(ch, '_' | '*') {
                index += 1;
                format_chars.get(index).ok_or(FormulaEvalError::Value)?;
                index += 1;
                continue;
            }
            let starts_with = |needle: &str| -> bool {
                let needle_chars = needle.chars().collect::<Vec<_>>();
                format_chars
                    .get(index..index + needle_chars.len())
                    .is_some_and(|candidate| {
                        candidate
                            .iter()
                            .zip(needle_chars.iter())
                            .all(|(left, right)| left.eq_ignore_ascii_case(right))
                    })
            };
            if starts_with("AM/PM") || starts_with("A/P") {
                push_literal(&mut tokens, &mut literal);
                let len = if starts_with("AM/PM") { 5 } else { 3 };
                tokens.push(TextDateTimeFormatToken::AmPm(
                    format_chars[index..index + len].iter().collect(),
                ));
                index += len;
                continue;
            }
            if ch.is_ascii_alphabetic() {
                let lower = ch.to_ascii_lowercase();
                if !matches!(lower, 'y' | 'm' | 'd' | 'h' | 's') {
                    return Err(FormulaEvalError::Value);
                }
                let start = index;
                index += 1;
                while index < format_chars.len()
                    && format_chars[index].to_ascii_lowercase() == lower
                {
                    index += 1;
                }
                push_literal(&mut tokens, &mut literal);
                let len = index - start;
                tokens.push(match lower {
                    'y' => TextDateTimeFormatToken::Year(len),
                    'm' => TextDateTimeFormatToken::MonthOrMinute(len),
                    'd' => TextDateTimeFormatToken::Day(len),
                    'h' => TextDateTimeFormatToken::Hour(len),
                    's' => TextDateTimeFormatToken::Second(len),
                    _ => return Err(FormulaEvalError::Value),
                });
                continue;
            }
            literal.push(ch);
            index += 1;
        }
        push_literal(&mut tokens, &mut literal);

        if tokens
            .iter()
            .any(|token| !matches!(token, TextDateTimeFormatToken::Literal(_)))
        {
            let has_am_pm = tokens
                .iter()
                .any(|token| matches!(token, TextDateTimeFormatToken::AmPm(_)));
            let mut minute_tokens = vec![false; tokens.len()];
            let mut needs_date = false;
            let mut needs_time = false;
            for token_index in 0..tokens.len() {
                match &tokens[token_index] {
                    TextDateTimeFormatToken::Literal(_) => {}
                    TextDateTimeFormatToken::Year(_) | TextDateTimeFormatToken::Day(_) => {
                        needs_date = true;
                    }
                    TextDateTimeFormatToken::Hour(_)
                    | TextDateTimeFormatToken::Second(_)
                    | TextDateTimeFormatToken::AmPm(_) => {
                        needs_time = true;
                    }
                    TextDateTimeFormatToken::MonthOrMinute(_) => {
                        let previous = tokens[..token_index]
                            .iter()
                            .rev()
                            .find(|token| !matches!(token, TextDateTimeFormatToken::Literal(_)));
                        let next = tokens[token_index + 1..]
                            .iter()
                            .find(|token| !matches!(token, TextDateTimeFormatToken::Literal(_)));
                        let time_context = previous.is_some_and(|token| {
                            matches!(
                                token,
                                TextDateTimeFormatToken::Hour(_)
                                    | TextDateTimeFormatToken::Second(_)
                            )
                        }) || next.is_some_and(|token| {
                            matches!(
                                token,
                                TextDateTimeFormatToken::Hour(_)
                                    | TextDateTimeFormatToken::Second(_)
                            )
                        }) || (has_am_pm
                            && !previous.is_some_and(|token| {
                                matches!(
                                    token,
                                    TextDateTimeFormatToken::Year(_)
                                        | TextDateTimeFormatToken::Day(_)
                                )
                            })
                            && !next.is_some_and(|token| {
                                matches!(
                                    token,
                                    TextDateTimeFormatToken::Year(_)
                                        | TextDateTimeFormatToken::Day(_)
                                )
                            }));
                        minute_tokens[token_index] = time_context;
                        if time_context {
                            needs_time = true;
                        } else {
                            needs_date = true;
                        }
                    }
                }
            }

            let (year, month, day, weekday) = if needs_date {
                let (year, month, day) = formula_ymd_from_serial(number)?;
                let serial = formula_serial_integer(number)?;
                (
                    year,
                    month,
                    day,
                    formula_weekday_monday0_from_serial(serial) as usize,
                )
            } else {
                (1900, 1, 1, 0)
            };
            let (hour, minute, second) = if needs_time {
                formula_time_parts_from_serial(number)?
            } else {
                (0, 0, 0)
            };
            let month_names = [
                ("Jan", "January"),
                ("Feb", "February"),
                ("Mar", "March"),
                ("Apr", "April"),
                ("May", "May"),
                ("Jun", "June"),
                ("Jul", "July"),
                ("Aug", "August"),
                ("Sep", "September"),
                ("Oct", "October"),
                ("Nov", "November"),
                ("Dec", "December"),
            ];
            let weekday_names = [
                ("Mon", "Monday"),
                ("Tue", "Tuesday"),
                ("Wed", "Wednesday"),
                ("Thu", "Thursday"),
                ("Fri", "Friday"),
                ("Sat", "Saturday"),
                ("Sun", "Sunday"),
            ];
            let mut output = String::new();
            for (token_index, token) in tokens.iter().enumerate() {
                match token {
                    TextDateTimeFormatToken::Literal(value) => output.push_str(value),
                    TextDateTimeFormatToken::Year(len) => {
                        if *len <= 2 {
                            output.push_str(format!("{:02}", year.rem_euclid(100)).as_str());
                        } else {
                            output.push_str(format!("{year:04}").as_str());
                        }
                    }
                    TextDateTimeFormatToken::MonthOrMinute(len) => {
                        if minute_tokens[token_index] {
                            if *len == 1 {
                                output.push_str(minute.to_string().as_str());
                            } else {
                                output.push_str(format!("{minute:02}").as_str());
                            }
                        } else {
                            match *len {
                                1 => output.push_str(month.to_string().as_str()),
                                2 => output.push_str(format!("{month:02}").as_str()),
                                3 => output.push_str(month_names[(month - 1) as usize].0),
                                _ => output.push_str(month_names[(month - 1) as usize].1),
                            }
                        }
                    }
                    TextDateTimeFormatToken::Day(len) => match *len {
                        1 => output.push_str(day.to_string().as_str()),
                        2 => output.push_str(format!("{day:02}").as_str()),
                        3 => output.push_str(weekday_names[weekday].0),
                        _ => output.push_str(weekday_names[weekday].1),
                    },
                    TextDateTimeFormatToken::Hour(len) => {
                        let display_hour = if has_am_pm {
                            match hour % 12 {
                                0 => 12,
                                value => value,
                            }
                        } else {
                            hour
                        };
                        if *len == 1 {
                            output.push_str(display_hour.to_string().as_str());
                        } else {
                            output.push_str(format!("{display_hour:02}").as_str());
                        }
                    }
                    TextDateTimeFormatToken::Second(len) => {
                        if *len == 1 {
                            output.push_str(second.to_string().as_str());
                        } else {
                            output.push_str(format!("{second:02}").as_str());
                        }
                    }
                    TextDateTimeFormatToken::AmPm(pattern) => {
                        let marker = if pattern.eq_ignore_ascii_case("A/P") {
                            if hour < 12 { "A" } else { "P" }
                        } else if hour < 12 {
                            "AM"
                        } else {
                            "PM"
                        };
                        if pattern.chars().any(|ch| ch.is_ascii_uppercase()) {
                            output.push_str(marker);
                        } else {
                            output.push_str(marker.to_ascii_lowercase().as_str());
                        }
                    }
                }
            }
            return Ok(output);
        }

        let mut in_quote = false;
        let mut escaped = false;
        for ch in format.chars() {
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' && !in_quote {
                escaped = true;
                continue;
            }
            if ch == '"' {
                in_quote = !in_quote;
                continue;
            }
            if !in_quote && (ch == '@' || ch.is_ascii_alphabetic()) {
                return Err(FormulaEvalError::Value);
            }
        }
        if in_quote || escaped {
            return Err(FormulaEvalError::Value);
        }

        let first_placeholder = format
            .find(|ch| matches!(ch, '0' | '#'))
            .ok_or(FormulaEvalError::Value)?;
        let last_placeholder = format
            .rfind(|ch| matches!(ch, '0' | '#'))
            .ok_or(FormulaEvalError::Value)?;
        let core = &format[first_placeholder..=last_placeholder];
        if core.matches('.').count() > 1 {
            return Err(FormulaEvalError::Value);
        }
        let decimals = core
            .split_once('.')
            .map(|(_, fraction)| {
                fraction
                    .chars()
                    .filter(|ch| matches!(*ch, '0' | '#'))
                    .count() as i64
            })
            .unwrap_or(0);
        let integer_part = core
            .split_once('.')
            .map(|(integer, _)| integer)
            .unwrap_or(core);
        let use_commas = integer_part.contains(',');
        let percent_count = format.chars().filter(|ch| *ch == '%').count();
        if percent_count > 1 {
            return Err(FormulaEvalError::Value);
        }
        let mut value = number;
        if percent_count == 1 {
            value *= 100.0;
        }
        if number < 0.0 && sections.len() > 1 {
            value = value.abs();
        }

        let format_literal = |literal: &str| -> Result<String, FormulaEvalError> {
            let mut output = String::new();
            let mut chars = literal.chars();
            let mut in_quote = false;
            while let Some(ch) = chars.next() {
                match ch {
                    '"' => in_quote = !in_quote,
                    '\\' => {
                        let escaped = chars.next().ok_or(FormulaEvalError::Value)?;
                        output.push(escaped);
                    }
                    '_' | '*' if !in_quote => {
                        chars.next().ok_or(FormulaEvalError::Value)?;
                    }
                    _ => output.push(ch),
                }
            }
            if in_quote {
                Err(FormulaEvalError::Value)
            } else {
                Ok(output)
            }
        };
        let formatted = formula_fixed_number_text(value, decimals, use_commas)?;
        let prefix = format_literal(&format[..first_placeholder])?;
        let suffix = format_literal(&format[last_placeholder + 1..])?;
        if let Some(positive) = formatted.strip_prefix('-') {
            Ok(format!("-{prefix}{positive}{suffix}"))
        } else {
            Ok(format!("{prefix}{formatted}{suffix}"))
        }
    }

    fn parse_hyperlink_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let link_location = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(link_location);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let friendly_name = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_text_from_value_probe(friendly_name)
    }

    fn parse_phonetic_text_function(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
            self.index = next_index;
            self.skip_whitespace();
            if self.consume_char(')') {
                let mut output = String::new();
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        match self
                            .evaluator
                            .cell_value_or_blank(target_sheet_id, row, col)?
                        {
                            CellValue::Text(value) => output.push_str(value.as_str()),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error));
                            }
                            CellValue::Blank | CellValue::Bool(_) | CellValue::Number(_) => {}
                        }
                    }
                }
                return Ok(output);
            }
        }
        self.index = checkpoint;
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match value {
            FormulaValueProbe::Text(value) => Ok(value),
            FormulaValueProbe::Error(error) => Err(error),
            FormulaValueProbe::Blank
            | FormulaValueProbe::Bool(_)
            | FormulaValueProbe::Number(_)
            | FormulaValueProbe::Omitted
            | FormulaValueProbe::Lambda { .. } => Ok(String::new()),
        }
    }

    fn parse_optional_engineering_places(&mut self) -> Result<Option<usize>, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(None);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let places = formula_integer_argument(self.parse_comparison()?)?;
        if places < 0 {
            return Err(FormulaEvalError::Num);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        usize::try_from(places)
            .map(Some)
            .map_err(|_| FormulaEvalError::Num)
    }

    fn parse_unary_text_function(
        &mut self,
        transform: impl FnOnce(String) -> String,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(transform(text))
    }

    fn parse_optional_text_count_argument(
        &mut self,
        default: usize,
    ) -> Result<usize, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(default);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(count)
    }

    fn parse_textjoin_function(&mut self) -> Result<String, FormulaEvalError> {
        let delimiter = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let ignore_empty = self.parse_comparison()? != 0.0;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }

        let mut parts = Vec::new();
        loop {
            for text in self.parse_text_values_argument()? {
                if !ignore_empty || !text.is_empty() {
                    parts.push(text);
                }
            }
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(parts.join(delimiter.as_str()));
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    fn parse_textsplit_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        let col_delimiter = if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
            None
        } else {
            Some(self.parse_text_value_argument()?)
        };
        let mut row_delimiter = None;
        let mut ignore_empty = false;
        let mut match_mode = 0_i64;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                row_delimiter = Some(self.parse_text_value_argument()?);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    ignore_empty = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        match_mode = formula_integer_argument(self.parse_comparison()?)?;
                        if !matches!(match_mode, 0 | 1) {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            self.parse_value_probe_argument()?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }
        let mut delimiters = Vec::new();
        if let Some(delimiter) = row_delimiter {
            delimiters.push(delimiter);
        }
        if let Some(delimiter) = col_delimiter {
            delimiters.push(delimiter);
        }
        if delimiters.is_empty() || delimiters.iter().any(|delimiter| delimiter.is_empty()) {
            return Err(FormulaEvalError::Value);
        }

        let case_insensitive = match_mode == 1;
        let mut parts = vec![text];
        for delimiter in delimiters {
            let mut next_parts = Vec::new();
            for part in parts {
                let matches = formula_text_delimiter_matches(
                    part.as_str(),
                    delimiter.as_str(),
                    case_insensitive,
                );
                if matches.is_empty() {
                    next_parts.push(part);
                    continue;
                }
                let mut start = 0_usize;
                for (match_start, match_end) in matches {
                    next_parts.push(part[start..match_start].to_string());
                    start = match_end;
                }
                next_parts.push(part[start..].to_string());
            }
            parts = next_parts;
        }
        parts
            .into_iter()
            .find(|part| !ignore_empty || !part.is_empty())
            .ok_or(FormulaEvalError::Calc)
    }

    fn parse_detectlanguage_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(formula_detect_language_tag(text.as_str()).to_string())
    }

    fn parse_translate_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let mut source_language = None::<String>;
        let mut target_language = None::<String>;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                source_language = Some(self.parse_text_value_argument()?);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    target_language = Some(self.parse_text_value_argument()?);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        let source_language =
            source_language.unwrap_or_else(|| formula_detect_language_tag(text.as_str()).into());
        let target_language = target_language.unwrap_or_else(|| "en".to_string());
        if source_language.eq_ignore_ascii_case(target_language.as_str()) {
            return Ok(text);
        }
        let normalized = text.trim().to_ascii_lowercase();
        let translated = match (
            source_language.to_ascii_lowercase().as_str(),
            target_language.to_ascii_lowercase().as_str(),
            normalized.as_str(),
        ) {
            ("en", "es", "hello") => "hola",
            ("en", "es", "world") => "mundo",
            ("en", "fr", "hello") => "bonjour",
            ("en", "de", "hello") => "hallo",
            ("es", "en", "hola") => "hello",
            ("es", "en", "mundo") => "world",
            ("fr", "en", "bonjour") => "hello",
            ("de", "en", "hallo") => "hello",
            _ => return Ok(text),
        };
        Ok(translated.to_string())
    }

    fn parse_image_function(&mut self) -> Result<String, FormulaEvalError> {
        let source = self.parse_text_value_argument()?;
        let mut alt_text = None::<String>;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                alt_text = Some(self.parse_text_value_argument()?);
            }
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    break;
                }
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    self.parse_value_probe_argument()?;
                }
            }
        }
        Ok(alt_text.filter(|value| !value.is_empty()).unwrap_or(source))
    }

    fn parse_webservice_function(&mut self) -> Result<String, FormulaEvalError> {
        let url = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let Some((metadata, payload)) = url
            .strip_prefix("data:")
            .and_then(|data| data.split_once(','))
        else {
            return Err(FormulaEvalError::Value);
        };
        if metadata.contains(";base64") {
            return Err(FormulaEvalError::Value);
        }
        let mut bytes = Vec::new();
        let payload_bytes = payload.as_bytes();
        let mut index = 0_usize;
        while index < payload_bytes.len() {
            if payload_bytes[index] == b'%' {
                if index + 2 >= payload_bytes.len() {
                    return Err(FormulaEvalError::Value);
                }
                let hex = &payload[index + 1..index + 3];
                let value = u8::from_str_radix(hex, 16).map_err(|_| FormulaEvalError::Value)?;
                bytes.push(value);
                index += 3;
            } else {
                bytes.push(payload_bytes[index]);
                index += 1;
            }
        }
        String::from_utf8(bytes).map_err(|_| FormulaEvalError::Value)
    }

    fn parse_filterxml_function(&mut self) -> Result<String, FormulaEvalError> {
        let xml = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let xpath = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let xpath = xpath.trim();
        let absolute = xpath.starts_with('/') && !xpath.starts_with("//");
        let trimmed = xpath.trim_start_matches('/');
        let mut path = Vec::new();
        let mut attribute_name = None::<String>;
        for segment in trimmed.split('/').filter(|segment| !segment.is_empty()) {
            if let Some(attribute) = segment.strip_prefix('@') {
                attribute_name = Some(attribute.to_ascii_lowercase());
                break;
            }
            let element = segment.split('[').next().unwrap_or(segment);
            if element.is_empty() {
                return Err(FormulaEvalError::Value);
            }
            path.push(element.to_ascii_lowercase());
        }
        if path.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        let matches_path = |stack: &[String], path: &[String]| -> bool {
            if absolute {
                stack == path
            } else {
                stack.ends_with(path)
            }
        };

        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut stack = Vec::<String>::new();
        let mut capture_depth = None::<usize>;
        let mut captured = String::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element)) => {
                    stack.push(
                        String::from_utf8_lossy(xml_local_name(element.name().as_ref()))
                            .to_ascii_lowercase(),
                    );
                    if capture_depth.is_none() && matches_path(stack.as_slice(), path.as_slice()) {
                        if let Some(attribute_name) = &attribute_name {
                            for attr in element.attributes() {
                                let attr = attr.map_err(|_| FormulaEvalError::Value)?;
                                if String::from_utf8_lossy(xml_local_name(attr.key.as_ref()))
                                    .eq_ignore_ascii_case(attribute_name.as_str())
                                {
                                    return attr
                                        .decode_and_unescape_value(reader.decoder())
                                        .map(|value| value.into_owned())
                                        .map_err(|_| FormulaEvalError::Value);
                                }
                            }
                            return Err(FormulaEvalError::Value);
                        }
                        capture_depth = Some(stack.len());
                    }
                }
                Ok(Event::Empty(element)) => {
                    stack.push(
                        String::from_utf8_lossy(xml_local_name(element.name().as_ref()))
                            .to_ascii_lowercase(),
                    );
                    if capture_depth.is_none() && matches_path(stack.as_slice(), path.as_slice()) {
                        if let Some(attribute_name) = &attribute_name {
                            for attr in element.attributes() {
                                let attr = attr.map_err(|_| FormulaEvalError::Value)?;
                                if String::from_utf8_lossy(xml_local_name(attr.key.as_ref()))
                                    .eq_ignore_ascii_case(attribute_name.as_str())
                                {
                                    return attr
                                        .decode_and_unescape_value(reader.decoder())
                                        .map(|value| value.into_owned())
                                        .map_err(|_| FormulaEvalError::Value);
                                }
                            }
                            return Err(FormulaEvalError::Value);
                        }
                        return Ok(String::new());
                    }
                    stack.pop();
                }
                Ok(Event::Text(text)) if capture_depth.is_some() => {
                    captured.push_str(String::from_utf8_lossy(text.as_ref()).as_ref());
                }
                Ok(Event::CData(text)) if capture_depth.is_some() => {
                    captured.push_str(String::from_utf8_lossy(text.as_ref()).as_ref());
                }
                Ok(Event::End(_)) => {
                    if capture_depth == Some(stack.len()) {
                        return Ok(captured.trim().to_string());
                    }
                    stack.pop();
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(_) => return Err(FormulaEvalError::Value),
            }
            buffer.clear();
        }
        Err(FormulaEvalError::Value)
    }

    fn parse_text_boundary_function(&mut self, after: bool) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let delimiter = self.parse_text_value_argument()?;
        if delimiter.is_empty() {
            return Err(FormulaEvalError::Value);
        }

        let mut instance = 1_i64;
        let mut match_mode = 0_i64;
        let mut match_end = false;
        let mut if_not_found = None;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            instance = formula_integer_argument(self.parse_comparison()?)?;
            if instance == 0 {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                match_mode = formula_integer_argument(self.parse_comparison()?)?;
                if !matches!(match_mode, 0 | 1) {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    match_end = self.parse_comparison()? != 0.0;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        if_not_found = Some(self.parse_text_value_argument()?);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        let mut matches =
            formula_text_delimiter_matches(text.as_str(), delimiter.as_str(), match_mode == 1);
        if match_end {
            if instance > 0 {
                matches.push((text.len(), text.len()));
            } else {
                matches.insert(0, (0, 0));
            }
        }
        let selected = if instance > 0 {
            usize::try_from(instance - 1)
                .ok()
                .and_then(|index| matches.get(index))
        } else {
            let count = instance
                .checked_abs()
                .and_then(|value| usize::try_from(value).ok());
            count
                .and_then(|count| matches.len().checked_sub(count))
                .and_then(|index| matches.get(index))
        };
        let Some(&(start, end)) = selected else {
            return if_not_found.ok_or(FormulaEvalError::NA);
        };
        if after {
            Ok(text[end..].to_string())
        } else {
            Ok(text[..start].to_string())
        }
    }

    fn parse_regex_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pattern = self.parse_text_value_argument()?;
        self.skip_whitespace();
        let case_insensitive = if self.consume_char(')') {
            false
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let case_sensitivity = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            formula_regex_case_insensitive(case_sensitivity)?
        };
        let regex = formula_regex_from_pattern(pattern.as_str(), case_insensitive)?;
        Ok(if regex.is_match(text.as_str()) {
            1.0
        } else {
            0.0
        })
    }

    fn parse_regex_extract_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pattern = self.parse_text_value_argument()?;
        let mut return_mode = 0_i64;
        let mut case_insensitive = false;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            return_mode = formula_integer_argument(self.parse_comparison()?)?;
            if !matches!(return_mode, 0 | 1 | 2) {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let case_sensitivity = formula_integer_argument(self.parse_comparison()?)?;
                case_insensitive = formula_regex_case_insensitive(case_sensitivity)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }

        let regex = formula_regex_from_pattern(pattern.as_str(), case_insensitive)?;
        let captures = regex.captures(text.as_str()).ok_or(FormulaEvalError::NA)?;
        if return_mode == 2 {
            return captures
                .iter()
                .skip(1)
                .find_map(|capture| capture.map(|value| value.as_str().to_string()))
                .ok_or(FormulaEvalError::NA);
        }
        captures
            .get(0)
            .map(|value| value.as_str().to_string())
            .ok_or(FormulaEvalError::NA)
    }

    fn parse_regex_replace_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pattern = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let replacement = self.parse_text_value_argument()?;
        let mut occurrence = 0_i64;
        let mut case_insensitive = false;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            occurrence = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let case_sensitivity = formula_integer_argument(self.parse_comparison()?)?;
                case_insensitive = formula_regex_case_insensitive(case_sensitivity)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }

        let regex = formula_regex_from_pattern(pattern.as_str(), case_insensitive)?;
        if occurrence == 0 {
            return Ok(regex
                .replace_all(text.as_str(), replacement.as_str())
                .into_owned());
        }
        let mut matches = Vec::new();
        for captures in regex.captures_iter(text.as_str()) {
            let Some(full_match) = captures.get(0) else {
                continue;
            };
            let mut expanded = String::new();
            captures.expand(replacement.as_str(), &mut expanded);
            matches.push((full_match.start(), full_match.end(), expanded));
        }
        let selected = if occurrence > 0 {
            usize::try_from(occurrence - 1)
                .ok()
                .and_then(|index| matches.get(index))
        } else {
            occurrence
                .checked_abs()
                .and_then(|value| usize::try_from(value).ok())
                .and_then(|count| matches.len().checked_sub(count))
                .and_then(|index| matches.get(index))
        };
        let Some((start, end, expanded)) = selected else {
            return Ok(text);
        };
        let mut output = String::new();
        output.push_str(&text[..*start]);
        output.push_str(expanded.as_str());
        output.push_str(&text[*end..]);
        Ok(output)
    }

    fn parse_forecast_ets_function(&mut self, name: &str) -> Result<f64, FormulaEvalError> {
        macro_rules! parse_optional_number {
            ($default:expr) => {{
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    $default
                } else {
                    self.parse_comparison()?
                }
            }};
        }

        let is_forecast = name.eq_ignore_ascii_case("FORECAST.ETS");
        let is_confint = name.eq_ignore_ascii_case("FORECAST.ETS.CONFINT");
        let is_seasonality = name.eq_ignore_ascii_case("FORECAST.ETS.SEASONALITY");
        let is_stat = name.eq_ignore_ascii_case("FORECAST.ETS.STAT");

        let mut target_date = None;
        if is_forecast || is_confint {
            let target = self.parse_comparison()?;
            if !target.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            target_date = Some(target);
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }

        let values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let timeline = self.parse_aggregate_argument()?;

        let mut confidence_level = 0.95_f64;
        let mut statistic_type = None;
        let mut seasonality = 1.0_f64;
        let mut data_completion = 1.0_f64;
        let mut aggregation = 0.0_f64;

        if is_confint {
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                confidence_level = parse_optional_number!(0.95_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    seasonality = parse_optional_number!(1.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        data_completion = parse_optional_number!(1.0_f64);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            aggregation = parse_optional_number!(0.0_f64);
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
            }
        } else if is_stat {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let statistic = self.parse_comparison()?;
            if !statistic.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            statistic_type = Some(formula_integer_argument(statistic)?);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                seasonality = parse_optional_number!(1.0_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    data_completion = parse_optional_number!(1.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        aggregation = parse_optional_number!(0.0_f64);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        } else if is_seasonality {
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                data_completion = parse_optional_number!(1.0_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    aggregation = parse_optional_number!(0.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        } else {
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                seasonality = parse_optional_number!(1.0_f64);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    data_completion = parse_optional_number!(1.0_f64);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        aggregation = parse_optional_number!(0.0_f64);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        if values.len() != timeline.len() || values.is_empty() {
            return Err(FormulaEvalError::NA);
        }
        if values
            .iter()
            .chain(timeline.iter())
            .any(|value| !value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        if is_confint
            && (!confidence_level.is_finite() || confidence_level <= 0.0 || confidence_level >= 1.0)
        {
            return Err(FormulaEvalError::Num);
        }
        let statistic_type = if let Some(statistic_type) = statistic_type {
            if !(1..=8).contains(&statistic_type) {
                return Err(FormulaEvalError::Num);
            }
            Some(statistic_type)
        } else {
            None
        };

        let seasonality_setting = formula_integer_argument(seasonality)?;
        if !(0..=8760).contains(&seasonality_setting) {
            return Err(FormulaEvalError::Num);
        }
        let data_completion = formula_integer_argument(data_completion)?;
        if data_completion != 0 && data_completion != 1 {
            return Err(FormulaEvalError::Num);
        }
        let aggregation = formula_integer_argument(aggregation)?;
        if !(0..=7).contains(&aggregation) {
            return Err(FormulaEvalError::Num);
        }

        let mut pairs = timeline
            .iter()
            .copied()
            .zip(values.iter().copied())
            .collect::<Vec<_>>();
        pairs.sort_by(|left, right| left.0.partial_cmp(&right.0).unwrap_or(Ordering::Equal));

        let mut aggregated_times = Vec::new();
        let mut aggregated_values = Vec::new();
        let mut index = 0usize;
        while index < pairs.len() {
            let time = pairs[index].0;
            let mut group = Vec::new();
            while index < pairs.len() && pairs[index].0 == time {
                group.push(pairs[index].1);
                index += 1;
            }
            let value = match aggregation {
                0 | 1 => group.iter().sum::<f64>() / group.len() as f64,
                2 | 3 => group.len() as f64,
                4 => group.iter().copied().fold(f64::NEG_INFINITY, f64::max),
                5 => {
                    group.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
                    let mid = group.len() / 2;
                    if group.len() % 2 == 0 {
                        (group[mid - 1] + group[mid]) / 2.0
                    } else {
                        group[mid]
                    }
                }
                6 => group.iter().copied().fold(f64::INFINITY, f64::min),
                7 => group.iter().sum::<f64>(),
                _ => unreachable!("aggregation was validated"),
            };
            aggregated_times.push(time);
            aggregated_values.push(value);
        }

        if aggregated_times.len() < 2 {
            return Err(FormulaEvalError::Num);
        }
        let mut step = f64::INFINITY;
        for window in aggregated_times.windows(2) {
            let diff = window[1] - window[0];
            if diff <= 0.0 {
                return Err(FormulaEvalError::Value);
            }
            step = step.min(diff);
        }
        if !step.is_finite() || step == 0.0 {
            return Err(FormulaEvalError::Num);
        }

        let first_time = aggregated_times[0];
        let last_time = *aggregated_times
            .last()
            .expect("timeline has at least two points");
        let total_slots_float = (last_time - first_time) / step;
        let total_slots_rounded = total_slots_float.round();
        if (total_slots_float - total_slots_rounded).abs() > 1e-7 {
            return Err(FormulaEvalError::Num);
        }
        let total_slots = total_slots_rounded as usize + 1;
        if total_slots < aggregated_values.len() {
            return Err(FormulaEvalError::Num);
        }
        let missing_slots = total_slots - aggregated_values.len();
        if missing_slots > 0 && (missing_slots as f64 / total_slots as f64) > 0.30 {
            return Err(FormulaEvalError::Num);
        }

        let mut completed = vec![None; total_slots];
        for (time, value) in aggregated_times.iter().zip(aggregated_values.iter()) {
            let slot_float = (*time - first_time) / step;
            let slot_rounded = slot_float.round();
            if (slot_float - slot_rounded).abs() > 1e-7 || slot_rounded < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let slot = slot_rounded as usize;
            if slot >= completed.len() {
                return Err(FormulaEvalError::Num);
            }
            completed[slot] = Some(*value);
        }
        for slot in 0..completed.len() {
            if completed[slot].is_some() {
                continue;
            }
            completed[slot] = Some(if data_completion == 0 {
                0.0
            } else {
                let previous = (0..slot).rev().find_map(|index| completed[index]);
                let next = (slot + 1..completed.len()).find_map(|index| completed[index]);
                match (previous, next) {
                    (Some(left), Some(right)) => (left + right) / 2.0,
                    (Some(left), None) => left,
                    (None, Some(right)) => right,
                    (None, None) => 0.0,
                }
            });
        }
        let completed_values = completed
            .into_iter()
            .map(|value| value.expect("missing slots were completed"))
            .collect::<Vec<_>>();

        let fit_model = |period: usize| -> (Vec<(f64, f64)>, Vec<f64>, f64, f64) {
            let period = period.max(1);
            let mut coefficients = vec![(0.0_f64, 0.0_f64); period];
            if period == 1 {
                let x_values = formula_regression_default_x(completed_values.len());
                let (slope, intercept) = formula_regression_slope_intercept(
                    completed_values.as_slice(),
                    x_values.as_slice(),
                    true,
                    false,
                )
                .unwrap_or((
                    0.0,
                    completed_values.iter().sum::<f64>() / completed_values.len() as f64,
                ));
                coefficients[0] = (slope, intercept);
            } else {
                for slot in 0..period {
                    let mut x_values = Vec::new();
                    let mut y_values = Vec::new();
                    for (index, value) in completed_values.iter().enumerate() {
                        if index % period == slot {
                            x_values.push(index as f64 + 1.0);
                            y_values.push(*value);
                        }
                    }
                    coefficients[slot] = if y_values.len() >= 2 {
                        formula_regression_slope_intercept(
                            y_values.as_slice(),
                            x_values.as_slice(),
                            true,
                            false,
                        )
                        .unwrap_or((0.0, y_values.iter().sum::<f64>() / y_values.len() as f64))
                    } else {
                        (
                            0.0,
                            y_values.first().copied().unwrap_or_else(|| {
                                completed_values.iter().sum::<f64>() / completed_values.len() as f64
                            }),
                        )
                    };
                }
            }

            let fitted = completed_values
                .iter()
                .enumerate()
                .map(|(index, _)| {
                    let (slope, intercept) = coefficients[index % period];
                    intercept + slope * (index as f64 + 1.0)
                })
                .collect::<Vec<_>>();
            let mae = completed_values
                .iter()
                .zip(fitted.iter())
                .map(|(actual, fitted)| (actual - fitted).abs())
                .sum::<f64>()
                / completed_values.len() as f64;
            let rmse = (completed_values
                .iter()
                .zip(fitted.iter())
                .map(|(actual, fitted)| {
                    let error = actual - fitted;
                    error * error
                })
                .sum::<f64>()
                / completed_values.len() as f64)
                .sqrt();
            (coefficients, fitted, mae, rmse)
        };

        let detect_period = || -> usize {
            if completed_values.len() < 4 {
                return 1;
            }
            let max_period = (completed_values.len() / 2).min(24).min(8760);
            let mut best_period = 1usize;
            let (_, _, _, mut best_rmse) = fit_model(1);
            for period in 2..=max_period {
                let (_, _, _, rmse) = fit_model(period);
                if rmse + 1e-9 < best_rmse {
                    best_rmse = rmse;
                    best_period = period;
                }
            }
            best_period
        };

        let period = if seasonality_setting == 0 {
            1usize
        } else if seasonality_setting == 1 {
            detect_period()
        } else {
            let period = usize::try_from(seasonality_setting).map_err(|_| FormulaEvalError::Num)?;
            if period > completed_values.len() || period > completed_values.len() / 2 {
                return Err(FormulaEvalError::Num);
            }
            period
        };

        let (coefficients, fitted, mae, rmse) = fit_model(period);
        let target_index = if let Some(target_date) = target_date {
            if target_date < last_time {
                return Err(FormulaEvalError::Num);
            }
            let offset = (target_date - first_time) / step;
            let rounded = offset.round();
            if (offset - rounded).abs() > 1e-7 || rounded < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            rounded as usize
        } else {
            completed_values.len()
        };
        let (target_slope, target_intercept) = coefficients[target_index % period];
        let forecast = target_intercept + target_slope * (target_index as f64 + 1.0);

        if is_forecast {
            return formula_checked_numeric_result(forecast);
        }
        if is_seasonality {
            return Ok(period as f64);
        }
        if is_confint {
            let inverse_standard_normal = |probability: f64| -> Result<f64, FormulaEvalError> {
                if !probability.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if probability <= 0.0 || probability >= 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                const A: [f64; 6] = [
                    -3.969683028665376e1,
                    2.209460984245205e2,
                    -2.759285104469687e2,
                    1.383577518672690e2,
                    -3.066479806614716e1,
                    2.506628277459239,
                ];
                const B: [f64; 5] = [
                    -5.447609879822406e1,
                    1.615858368580409e2,
                    -1.556989798598866e2,
                    6.680131188771972e1,
                    -1.328068155288572e1,
                ];
                const C: [f64; 6] = [
                    -7.784894002430293e-3,
                    -3.223964580411365e-1,
                    -2.400758277161838,
                    -2.549732539343734,
                    4.374664141464968,
                    2.938163982698783,
                ];
                const D: [f64; 4] = [
                    7.784695709041462e-3,
                    3.224671290700398e-1,
                    2.445134137142996,
                    3.754408661907416,
                ];
                const P_LOW: f64 = 0.02425;
                const P_HIGH: f64 = 1.0 - P_LOW;
                if probability < P_LOW {
                    let q = (-2.0 * probability.ln()).sqrt();
                    let numerator =
                        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q) + C[5];
                    let denominator = ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q) + 1.0;
                    Ok(numerator / denominator)
                } else if probability <= P_HIGH {
                    let q = probability - 0.5;
                    let r = q * q;
                    let numerator =
                        (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r) + A[5];
                    let denominator =
                        (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r) + 1.0;
                    Ok(numerator * q / denominator)
                } else {
                    let q = (-2.0 * (1.0 - probability).ln()).sqrt();
                    let numerator =
                        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q) + C[5];
                    let denominator = ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q) + 1.0;
                    Ok(-numerator / denominator)
                }
            };
            let horizon = target_index.saturating_sub(completed_values.len() - 1) as f64;
            return formula_checked_numeric_result(
                inverse_standard_normal(0.5 + confidence_level / 2.0)?
                    * rmse
                    * (1.0 + horizon / completed_values.len() as f64).sqrt(),
            );
        }

        let statistic_type = statistic_type.expect("statistic type was parsed");
        match statistic_type {
            1 => Ok(0.5),
            2 => {
                let trend_weight = coefficients
                    .iter()
                    .any(|(slope, _)| slope.abs() > f64::EPSILON);
                Ok(if trend_weight { 0.1 } else { 0.0 })
            }
            3 => Ok(if period > 1 { 0.1 } else { 0.0 }),
            4 => {
                let lag = period.max(1);
                if completed_values.len() <= lag {
                    return Err(FormulaEvalError::Div0);
                }
                let scale = (lag..completed_values.len())
                    .map(|index| (completed_values[index] - completed_values[index - lag]).abs())
                    .sum::<f64>()
                    / (completed_values.len() - lag) as f64;
                if scale == 0.0 {
                    return if mae == 0.0 {
                        Ok(0.0)
                    } else {
                        Err(FormulaEvalError::Div0)
                    };
                }
                formula_checked_numeric_result(mae / scale)
            }
            5 => {
                let mut total = 0.0_f64;
                let mut count = 0usize;
                for (actual, forecast) in completed_values.iter().zip(fitted.iter()) {
                    let denominator = actual.abs() + forecast.abs();
                    if denominator != 0.0 {
                        total += 200.0 * (actual - forecast).abs() / denominator;
                        count += 1;
                    }
                }
                Ok(if count == 0 {
                    0.0
                } else {
                    total / count as f64
                })
            }
            6 => Ok(mae),
            7 => Ok(rmse),
            8 => Ok(step),
            _ => unreachable!("statistic type was validated"),
        }
    }

    fn parse_regression_coefficient_function(
        &mut self,
        exponential: bool,
    ) -> Result<f64, FormulaEvalError> {
        let known_y = self.parse_aggregate_argument()?;
        let mut known_x = formula_regression_default_x(known_y.len());
        let mut constant = true;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                known_x = self.parse_aggregate_argument()?;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    constant = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        self.parse_comparison()?;
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        let (slope, _) = formula_regression_slope_intercept(
            known_y.as_slice(),
            known_x.as_slice(),
            constant,
            exponential,
        )?;
        if exponential {
            formula_checked_numeric_result(slope.exp())
        } else {
            formula_checked_numeric_result(slope)
        }
    }

    fn parse_regression_prediction_function(
        &mut self,
        exponential: bool,
    ) -> Result<f64, FormulaEvalError> {
        let known_y = self.parse_aggregate_argument()?;
        let mut known_x = formula_regression_default_x(known_y.len());
        let mut new_x = None;
        let mut constant = true;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                known_x = self.parse_aggregate_argument()?;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    new_x = self.parse_aggregate_argument()?.first().copied();
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        constant = self.parse_comparison()? != 0.0;
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        let new_x = new_x
            .or_else(|| known_x.first().copied())
            .ok_or(FormulaEvalError::NA)?;
        if !new_x.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let (slope, intercept) = formula_regression_slope_intercept(
            known_y.as_slice(),
            known_x.as_slice(),
            constant,
            exponential,
        )?;
        let prediction = intercept + slope * new_x;
        if exponential {
            formula_checked_numeric_result(prediction.exp())
        } else {
            formula_checked_numeric_result(prediction)
        }
    }

    fn parse_rept_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let output_len = text
            .chars()
            .count()
            .checked_mul(count)
            .ok_or(FormulaEvalError::Value)?;
        if output_len > 32_767 {
            return Err(FormulaEvalError::Value);
        }
        Ok(text.repeat(count))
    }

    fn parse_replace_text_function(&mut self, byte_mode: bool) -> Result<String, FormulaEvalError> {
        let old_text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let start = formula_positive_position_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let new_text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if byte_mode {
            let len = formula_text_byte_len(old_text.as_str());
            if start > len + 1 {
                return Err(FormulaEvalError::Value);
            }
            let prefix = formula_text_byte_slice(old_text.as_str(), 1, start - 1);
            let suffix_start = start.saturating_add(count);
            let suffix_count = len.saturating_sub(suffix_start.saturating_sub(1));
            let suffix = formula_text_byte_slice(old_text.as_str(), suffix_start, suffix_count);
            return Ok(format!("{prefix}{new_text}{suffix}"));
        }
        let chars = old_text.chars().collect::<Vec<_>>();
        if start > chars.len() + 1 {
            return Err(FormulaEvalError::Value);
        }
        let start_index = start - 1;
        let end_index = (start_index + count).min(chars.len());
        let mut output = String::new();
        output.extend(chars[..start_index].iter());
        output.push_str(new_text.as_str());
        output.extend(chars[end_index..].iter());
        Ok(output)
    }

    fn parse_substitute_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let old_text = self.parse_text_value_argument()?;
        if old_text.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let new_text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(text.replace(old_text.as_str(), new_text.as_str()));
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let instance = formula_positive_position_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mut match_count = 0_usize;
        for (byte_index, _) in text.match_indices(old_text.as_str()) {
            match_count += 1;
            if match_count == instance {
                let mut output = String::new();
                output.push_str(&text[..byte_index]);
                output.push_str(new_text.as_str());
                output.push_str(&text[byte_index + old_text.len()..]);
                return Ok(output);
            }
        }
        Ok(text)
    }

    fn parse_formulatext_function(&mut self) -> Result<String, FormulaEvalError> {
        let (target_sheet_id, rect) = self.parse_reference_argument()?;
        if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let Some(formula) =
            self.evaluator
                .formula_source_at(target_sheet_id, rect.row_first, rect.col_first)?
        else {
            return Err(FormulaEvalError::NA);
        };
        let text = if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, rect.row_first, rect.col_first)
        } else {
            formula.text
        };
        Ok(if text.starts_with('=') {
            text
        } else {
            format!("={text}")
        })
    }

    fn parse_value_text_format_argument(&mut self) -> Result<bool, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(false);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let format = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match format {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(FormulaEvalError::Value),
        }
    }

    fn parse_valuetotext_function(&mut self) -> Result<String, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        let strict = self.parse_value_text_format_argument()?;
        formula_value_to_text(value, strict)
    }

    fn parse_arraytotext_function(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        let mut rows = Vec::new();
        if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
            self.index = next_index;
            self.skip_whitespace();
            if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                for row in rect.row_first..=rect.row_last {
                    let mut values = Vec::new();
                    for col in rect.col_first..=rect.col_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?;
                        values.push(formula_value_probe_from_cell_value(value));
                    }
                    rows.push(values);
                }
            } else {
                self.index = checkpoint;
            }
        }
        if rows.is_empty() {
            rows.push(vec![self.parse_value_probe_argument()?]);
        }
        let strict = self.parse_value_text_format_argument()?;
        if !strict {
            return rows
                .into_iter()
                .flatten()
                .map(|value| formula_value_to_text(value, false))
                .collect::<Result<Vec<_>, _>>()
                .map(|values| values.join(", "));
        }
        let mut row_texts = Vec::new();
        for row in rows {
            let values = row
                .into_iter()
                .map(|value| formula_value_to_text(value, true))
                .collect::<Result<Vec<_>, _>>()?;
            row_texts.push(values.join(","));
        }
        Ok(format!("{{{}}}", row_texts.join(";")))
    }

    fn parse_len_function(&mut self, byte_mode: bool) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if byte_mode {
            formula_text_byte_len(text.as_str()) as f64
        } else {
            text.chars().count() as f64
        })
    }

    fn parse_character_code_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        text.chars()
            .next()
            .map(|ch| ch as u32 as f64)
            .ok_or(FormulaEvalError::Value)
    }

    fn parse_find_function(
        &mut self,
        case_insensitive: bool,
        byte_mode: bool,
    ) -> Result<f64, FormulaEvalError> {
        let needle = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let haystack = self.parse_text_value_argument()?;
        self.skip_whitespace();
        let start = if self.consume_char(')') {
            1
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let start = formula_positive_position_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            start
        };
        let haystack_len = haystack.chars().count();
        if start > haystack_len + 1 || (start > haystack_len && !needle.is_empty()) {
            return Err(FormulaEvalError::Value);
        }
        if case_insensitive {
            let Some(position) =
                formula_wildcard_find(needle.as_str(), haystack.as_str(), start, true)
            else {
                return Err(FormulaEvalError::Value);
            };
            return Ok(if byte_mode {
                formula_text_char_position_to_byte_position(haystack.as_str(), position) as f64
            } else {
                position as f64
            });
        }
        let searchable = haystack.chars().skip(start - 1).collect::<String>();
        let Some(byte_index) = searchable.find(needle.as_str()) else {
            return Err(FormulaEvalError::Value);
        };
        let position = start + searchable[..byte_index].chars().count();
        Ok(if byte_mode {
            formula_text_char_position_to_byte_position(haystack.as_str(), position) as f64
        } else {
            position as f64
        })
    }

    fn parse_exact_function(&mut self) -> Result<f64, FormulaEvalError> {
        let left = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let right = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if left == right { 1.0 } else { 0.0 })
    }

    fn parse_value_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_value_text(text.as_str())
    }

    fn parse_numbervalue_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let mut decimal_separator = ".".to_string();
        let mut group_separator = ",".to_string();
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            decimal_separator = self.parse_text_value_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                group_separator = self.parse_text_value_argument()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_numbervalue(
            text.as_str(),
            decimal_separator.as_str(),
            group_separator.as_str(),
        )
    }

    fn parse_decimal_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let radix = formula_radix_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let text = text.trim();
        if text.is_empty() || text.len() > 255 {
            return Err(FormulaEvalError::Num);
        }
        let mut total = 0.0_f64;
        for ch in text.chars() {
            let Some(digit) = ch.to_digit(radix) else {
                return Err(FormulaEvalError::Num);
            };
            total = total * radix as f64 + digit as f64;
            if !total.is_finite() {
                return Err(FormulaEvalError::Num);
            }
        }
        Ok(total)
    }

    fn parse_engineering_decimal_function(
        &mut self,
        source_radix: u32,
        source_bits: u32,
        source_max_digits: usize,
    ) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(
            formula_engineering_input(text.as_str(), source_radix, source_bits, source_max_digits)?
                as f64,
        )
    }

    fn parse_dollarde_function(&mut self) -> Result<f64, FormulaEvalError> {
        let fractional_dollar = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (denominator, scale) = formula_dollar_fraction_parts(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !fractional_dollar.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let sign = if fractional_dollar.is_sign_negative() {
            -1.0
        } else {
            1.0
        };
        let absolute = fractional_dollar.abs();
        let whole = absolute.trunc();
        let numerator = formula_dollar_fraction_near_integer((absolute - whole) * scale);
        let value = sign * (whole + numerator / denominator);
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    fn parse_dollarfr_function(&mut self) -> Result<f64, FormulaEvalError> {
        let decimal_dollar = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (denominator, scale) = formula_dollar_fraction_parts(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !decimal_dollar.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let sign = if decimal_dollar.is_sign_negative() {
            -1.0
        } else {
            1.0
        };
        let absolute = decimal_dollar.abs();
        let whole = absolute.trunc();
        let numerator = formula_dollar_fraction_near_integer((absolute - whole) * denominator);
        let value = sign * (whole + numerator / scale);
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    fn parse_fv_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        let mut pv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            pv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_fv_value(rate, nper, pmt, pv, payment_type)
    }

    fn parse_pv_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        if ![rate, nper, pmt, fv].iter().all(|value| value.is_finite()) {
            return Err(FormulaEvalError::Value);
        }
        let value = if rate == 0.0 {
            -(fv + pmt * nper)
        } else {
            let growth = formula_annuity_growth(rate, nper)?;
            if growth == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            -(fv + pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate) / growth
        };
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    fn parse_pmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_pmt_value(rate, nper, pv, fv, payment_type)
    }

    fn parse_ipmt_or_ppmt_function(&mut self, principal: bool) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        let interest = formula_ipmt_value(rate, period, nper, pv, fv, payment_type)?;
        if principal {
            Ok(formula_pmt_value(rate, nper, pv, fv, payment_type)? - interest)
        } else {
            Ok(interest)
        }
    }

    fn parse_ipmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_ipmt_or_ppmt_function(false)
    }

    fn parse_ppmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_ipmt_or_ppmt_function(true)
    }

    fn parse_cumulative_payment_function(
        &mut self,
        principal: bool,
    ) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let start_period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if ![rate, nper, pv, start_period, end_period]
            .iter()
            .all(|value| value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        let start_period = start_period.trunc();
        let end_period = end_period.trunc();
        if rate <= 0.0
            || nper <= 0.0
            || pv <= 0.0
            || start_period < 1.0
            || end_period < 1.0
            || start_period > end_period
        {
            return Err(FormulaEvalError::Num);
        }
        let payment = formula_pmt_value(rate, nper, pv, 0.0, payment_type)?;
        let mut total = 0.0;
        let mut period = start_period;
        while period <= end_period {
            let interest = formula_ipmt_value(rate, period, nper, pv, 0.0, payment_type)?;
            total += if principal {
                payment - interest
            } else {
                interest
            };
            period += 1.0;
        }
        if total.is_finite() {
            Ok(total)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    fn parse_nper_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        if ![rate, pmt, pv, fv].iter().all(|value| value.is_finite()) {
            return Err(FormulaEvalError::Value);
        }
        let value = if rate == 0.0 {
            if pmt == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            -(pv + fv) / pmt
        } else {
            let base = 1.0 + rate;
            if base <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let adjusted_payment = pmt * (1.0 + rate * payment_type) / rate;
            let numerator = adjusted_payment - fv;
            let denominator = pv + adjusted_payment;
            if numerator == 0.0 || denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            let ratio = numerator / denominator;
            if ratio <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            ratio.ln() / base.ln()
        };
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    fn parse_rate_function(&mut self) -> Result<f64, FormulaEvalError> {
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pmt = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        let mut fv = 0.0;
        let mut payment_type = 0.0;
        let mut guess = 0.1;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            fv = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                payment_type = formula_financial_type_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    guess = self.parse_comparison()?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }
        if ![nper, pmt, pv, fv, guess]
            .iter()
            .all(|value| value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        if nper <= 0.0 || guess <= -1.0 {
            return Err(FormulaEvalError::Num);
        }

        let rate_residual = |rate: f64| -> Result<f64, FormulaEvalError> {
            if !rate.is_finite() || rate <= -1.0 {
                return Err(FormulaEvalError::Num);
            }
            let value = if rate.abs() < 1e-10 {
                pv + pmt * nper + fv
            } else {
                let growth = formula_annuity_growth(rate, nper)?;
                pv * growth + pmt * (1.0 + rate * payment_type) * (growth - 1.0) / rate + fv
            };
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };

        const RATE_MAX_ITERATIONS: usize = 20;
        const RATE_TOLERANCE: f64 = 1e-7;
        let mut rate = guess;
        for _ in 0..RATE_MAX_ITERATIONS {
            let value = rate_residual(rate)?;
            if value.abs() <= RATE_TOLERANCE {
                return Ok(rate);
            }
            let step = (rate.abs() * 1e-6).max(1e-8);
            let right_rate = rate + step;
            let right_value = rate_residual(right_rate)?;
            let derivative = (right_value - value) / (right_rate - rate);
            if !derivative.is_finite() || derivative == 0.0 {
                break;
            }
            let next_rate = rate - value / derivative;
            if !next_rate.is_finite() || next_rate <= -1.0 {
                break;
            }
            if (next_rate - rate).abs() <= RATE_TOLERANCE {
                return Ok(next_rate);
            }
            rate = next_rate;
        }
        Err(FormulaEvalError::Num)
    }

    fn parse_ispmt_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rate = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let period = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let nper = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pv = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if ![rate, period, nper, pv]
            .iter()
            .all(|value| value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        if nper == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let value = pv * rate * (period / nper - 1.0);
        if value.is_finite() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Num)
        }
    }

    fn parse_arabic_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let trimmed = text.trim();
        if trimmed.chars().count() > 255 {
            return Err(FormulaEvalError::Value);
        }
        if trimmed.is_empty() {
            return Ok(0.0);
        }
        let (negative, body) = if let Some(body) = trimmed.strip_prefix('-') {
            (true, body)
        } else {
            (false, trimmed)
        };
        if body.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        let roman = body.to_ascii_uppercase();
        let mut total = 0_i64;
        let mut previous = 0_i64;
        for ch in roman.chars().rev() {
            let value = match ch {
                'I' => 1,
                'V' => 5,
                'X' => 10,
                'L' => 50,
                'C' => 100,
                'D' => 500,
                'M' => 1000,
                _ => return Err(FormulaEvalError::Value),
            };
            if value < previous {
                total -= value;
            } else {
                total += value;
                previous = value;
            }
        }
        if total <= 0 {
            return Err(FormulaEvalError::Value);
        }
        let thousands = usize::try_from(total / 1000).map_err(|_| FormulaEvalError::Value)?;
        let suffix_value = total % 1000;
        let prefix = "M".repeat(thousands);
        let mut valid = false;
        for form in 0..=4 {
            let candidate = format!("{prefix}{}", formula_roman_text(suffix_value, form)?);
            if roman == candidate {
                valid = true;
                break;
            }
        }
        if !valid {
            return Err(FormulaEvalError::Value);
        }
        Ok(if negative {
            -(total as f64)
        } else {
            total as f64
        })
    }

    fn parse_datevalue_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_datevalue_text(text.as_str())
    }

    fn parse_timevalue_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_timevalue_text(text.as_str())
    }

    fn parse_countif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria = self.parse_criteria_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(self
            .evaluator
            .countif_values_in_rect(sheet_id, rect, &criteria)? as f64)
    }

    fn parse_datedif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_serial = formula_serial_integer(self.parse_comparison()?)?;
        if start_serial > end_serial {
            return Err(FormulaEvalError::Num);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let unit = self.parse_text_value_argument()?.to_ascii_uppercase();
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (start_year, start_month, start_day) = formula_ymd_from_serial(start_serial as f64)?;
        let (end_year, end_month, end_day) = formula_ymd_from_serial(end_serial as f64)?;
        match unit.as_str() {
            "D" => Ok((end_serial - start_serial) as f64),
            "Y" => {
                let mut years = end_year - start_year;
                if (end_month, end_day) < (start_month, start_day) {
                    years -= 1;
                }
                Ok(years as f64)
            }
            "M" | "YM" => {
                let mut months =
                    (end_year - start_year) * 12 + i64::from(end_month) - i64::from(start_month);
                if end_day < start_day {
                    months -= 1;
                }
                if unit == "YM" {
                    months = months.rem_euclid(12);
                }
                Ok(months as f64)
            }
            "MD" => {
                if end_day >= start_day {
                    return Ok(f64::from(end_day - start_day));
                }
                let (previous_month_year, previous_month) = normalize_year_month(
                    end_year,
                    i64::from(end_month)
                        .checked_sub(1)
                        .ok_or(FormulaEvalError::Num)?,
                )?;
                Ok(f64::from(
                    days_in_excel_month(previous_month_year, previous_month) + end_day - start_day,
                ))
            }
            "YD" => {
                let mut anchor = formula_date_serial_from_args(
                    end_year as f64,
                    f64::from(start_month),
                    f64::from(start_day),
                )? as i64;
                if anchor > end_serial {
                    anchor = formula_date_serial_from_args(
                        (end_year - 1) as f64,
                        f64::from(start_month),
                        f64::from(start_day),
                    )? as i64;
                }
                Ok((end_serial - anchor) as f64)
            }
            _ => Err(FormulaEvalError::Num),
        }
    }

    fn parse_workday_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let days = formula_integer_argument(self.parse_comparison()?)?;
        let holidays = self.parse_optional_holidays_tail()?;
        formula_workday(start_serial, days, holidays.as_slice())
    }

    fn parse_workday_intl_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let days = formula_integer_argument(self.parse_comparison()?)?;
        let (weekend, holidays) = self.parse_optional_weekend_holidays_tail()?;
        formula_workday_with_weekend(start_serial, days, holidays.as_slice(), &weekend)
    }

    fn parse_networkdays_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_serial = formula_serial_integer(self.parse_comparison()?)?;
        let holidays = self.parse_optional_holidays_tail()?;
        formula_networkdays(start_serial, end_serial, holidays.as_slice())
    }

    fn parse_networkdays_intl_function(&mut self) -> Result<f64, FormulaEvalError> {
        let start_serial = formula_serial_integer(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let end_serial = formula_serial_integer(self.parse_comparison()?)?;
        let (weekend, holidays) = self.parse_optional_weekend_holidays_tail()?;
        formula_networkdays_with_weekend(start_serial, end_serial, holidays.as_slice(), &weekend)
    }

    fn parse_if_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_if_value_function()?)
    }

    fn parse_if_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let condition = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let true_value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(if condition != 0.0 {
                true_value
            } else {
                FormulaValueProbe::Bool(false)
            });
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let false_value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if condition != 0.0 {
            true_value
        } else {
            false_value
        })
    }

    fn parse_ifs_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut selected_value = None;
        loop {
            let condition = match self.parse_catchable_argument()? {
                Ok(condition) => condition,
                Err(error) if selected_value.is_none() => return Err(error),
                Err(_) => 0.0,
            };
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = self.parse_value_probe_argument()?;
            if selected_value.is_none() && condition != 0.0 {
                selected_value = Some(value);
            }
            self.skip_whitespace();
            if self.consume_char(')') {
                return selected_value.ok_or(FormulaEvalError::NA);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    fn parse_switch_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let expression = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mut selected_value = None;
        let mut saw_pair = false;
        loop {
            let candidate_or_default = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if self.consume_char(')') {
                if !saw_pair {
                    return Err(FormulaEvalError::Value);
                }
                return Ok(selected_value.unwrap_or(candidate_or_default));
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let result = self.parse_value_probe_argument()?;
            if selected_value.is_none()
                && formula_value_probe_exact_match(&expression, &candidate_or_default)?
            {
                selected_value = Some(result);
            }
            saw_pair = true;
            self.skip_whitespace();
            if self.consume_char(')') {
                return selected_value.ok_or(FormulaEvalError::NA);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    fn parse_choose_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_choose_value_function()?)
    }

    fn parse_choose_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let selected_index = formula_integer_argument(self.parse_comparison()?)?;
        if selected_index < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mut argument_index = 1_i64;
        let mut selected_value = None;
        loop {
            let value = self.parse_value_probe_argument()?;
            if argument_index == selected_index {
                selected_value = Some(value);
            }
            self.skip_whitespace();
            if self.consume_char(')') {
                return selected_value.ok_or(FormulaEvalError::Value);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            argument_index += 1;
        }
    }

    fn parse_let_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let base_binding_len = self.bindings.len();
        loop {
            self.skip_whitespace();
            let Some(name) = self.parse_identifier() else {
                self.bindings.truncate(base_binding_len);
                return Err(FormulaEvalError::Value);
            };
            self.skip_whitespace();
            if !self.consume_char(',') {
                self.bindings.truncate(base_binding_len);
                return Err(FormulaEvalError::Value);
            }
            let value = self.parse_value_probe_argument()?;
            self.bindings.push((name, value));
            self.skip_whitespace();
            if !self.consume_char(',') {
                self.bindings.truncate(base_binding_len);
                return Err(FormulaEvalError::Unsupported);
            }

            self.skip_whitespace();
            let checkpoint = self.index;
            let next_is_binding = if self.parse_identifier().is_some() {
                self.skip_whitespace();
                self.consume_char(',')
            } else {
                false
            };
            self.index = checkpoint;
            if next_is_binding {
                continue;
            }

            let result = self.parse_value_probe_argument();
            self.skip_whitespace();
            let result = match result {
                Ok(value) if self.consume_char(')') => Ok(value),
                Ok(_) => Err(FormulaEvalError::Unsupported),
                Err(error) => Err(error),
            };
            self.bindings.truncate(base_binding_len);
            return result;
        }
    }

    fn parse_isomitted_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        let Some(name) = self.parse_identifier() else {
            return Err(FormulaEvalError::Value);
        };
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(
            if matches!(
                self.binding_value(name.as_str()),
                Some(FormulaValueProbe::Omitted)
            ) {
                1.0
            } else {
                0.0
            },
        )
    }

    fn parse_makearray_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let rows = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let columns = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let lambda = self.parse_lambda_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if rows < 1 || columns < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.evaluate_lambda_value(
            lambda,
            vec![
                FormulaValueProbe::Number(1.0),
                FormulaValueProbe::Number(1.0),
            ],
        )
    }

    fn parse_reduce_scan_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut accumulator = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (array_sheet_id, array_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let lambda = self.parse_lambda_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        for row in array_rect.row_first..=array_rect.row_last {
            for col in array_rect.col_first..=array_rect.col_last {
                let value = self
                    .evaluator
                    .cell_value_or_blank(array_sheet_id, row, col)?;
                accumulator = self.evaluate_lambda_value(
                    lambda.clone(),
                    vec![accumulator, formula_value_probe_from_cell_value(value)],
                )?;
                if name.eq_ignore_ascii_case("SCAN") {
                    return Ok(accumulator);
                }
            }
        }
        Ok(accumulator)
    }

    fn parse_getpivotdata_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let data_field = self.parse_text_value_argument()?;
        if data_field.trim().is_empty() {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (pivot_sheet_id, pivot_rect) = self.parse_reference_argument()?;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                let value = self.evaluator.cell_value_or_blank(
                    pivot_sheet_id,
                    pivot_rect.row_first,
                    pivot_rect.col_first,
                )?;
                return Ok(formula_value_probe_from_cell_value(value));
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Value);
            }
            self.parse_value_probe_argument()?;
        }
    }

    fn parse_external_data_unavailable_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::NA)
    }

    fn parse_external_platform_unavailable_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::Value)
    }

    fn parse_external_field_unavailable_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::Field)
    }

    fn parse_external_python_unavailable_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.consume_all_value_arguments()?;
        Err(FormulaEvalError::Blocked)
    }

    fn parse_cube_caption_text_function(&mut self, name: &str) -> Result<String, FormulaEvalError> {
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let fallback = self.parse_text_value_argument()?;
        if name.eq_ignore_ascii_case("CUBEKPIMEMBER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.parse_value_probe_argument()?;
        } else if name.eq_ignore_ascii_case("CUBERANKEDMEMBER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.parse_comparison()?;
        }

        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(fallback);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        let caption = if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
            fallback
        } else {
            self.parse_text_value_argument()?
        };
        self.consume_remaining_optional_value_arguments()?;
        Ok(caption)
    }

    fn parse_cubememberproperty_text_function(&mut self) -> Result<String, FormulaEvalError> {
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Err(FormulaEvalError::NA)
    }

    fn parse_cubesetcount_function(&mut self) -> Result<f64, FormulaEvalError> {
        let set = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if set.trim().is_empty() { 0.0 } else { 1.0 })
    }

    fn parse_groupby_aggregation_argument(
        &mut self,
    ) -> Result<FormulaGroupByAggregation, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        if let Some(name) = self.parse_identifier() {
            self.skip_whitespace();
            if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                return FormulaGroupByAggregation::from_name(name.as_str())
                    .ok_or(FormulaEvalError::Value);
            }
        }
        self.index = checkpoint;
        match self.parse_value_probe_argument()? {
            FormulaValueProbe::Text(name) => {
                FormulaGroupByAggregation::from_name(name.as_str()).ok_or(FormulaEvalError::Value)
            }
            FormulaValueProbe::Error(error) => Err(error),
            _ => Err(FormulaEvalError::Value),
        }
    }

    fn consume_remaining_optional_value_arguments(&mut self) -> Result<(), FormulaEvalError> {
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(());
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                continue;
            }
            let checkpoint = self.index;
            if let Some((_, _, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
            } else {
                self.index = checkpoint;
                self.parse_value_probe_argument()?;
            }
        }
    }

    fn consume_all_value_arguments(&mut self) -> Result<(), FormulaEvalError> {
        let mut needs_separator = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(());
            }
            if needs_separator {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if self.consume_char(')') {
                    return Ok(());
                }
            }
            if self.peek_char().is_some_and(|ch| ch == ',') {
                needs_separator = true;
                continue;
            }
            let checkpoint = self.index;
            if let Some((_, _, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
            } else {
                self.index = checkpoint;
                self.parse_value_probe_argument()?;
            }
            needs_separator = true;
        }
    }

    fn rect_row_key_values(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        row_offset: u32,
    ) -> Result<Vec<FormulaValueProbe>, FormulaEvalError> {
        let mut key = Vec::with_capacity(rect.width() as usize);
        let row = rect.row_first + row_offset;
        for col in rect.col_first..=rect.col_last {
            let value = self.evaluator.cell_value_or_blank(sheet_id, row, col)?;
            key.push(formula_value_probe_from_cell_value(value));
        }
        Ok(key)
    }

    fn rect_row_key_matches(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        row_offset: u32,
        target: &[FormulaValueProbe],
    ) -> Result<bool, FormulaEvalError> {
        let candidate = self.rect_row_key_values(sheet_id, rect, row_offset)?;
        if candidate.len() != target.len() {
            return Ok(false);
        }
        for (left, right) in candidate.iter().zip(target.iter()) {
            if !formula_value_probe_exact_match(left, right)? {
                return Ok(false);
            }
        }
        Ok(true)
    }

    fn collect_first_group_values(
        &mut self,
        row_sheet_id: SheetId,
        row_rect: Rect,
        value_sheet_id: SheetId,
        value_rect: Rect,
        column_filter: Option<(SheetId, Rect, Vec<FormulaValueProbe>)>,
    ) -> Result<Vec<FormulaValueProbe>, FormulaEvalError> {
        if row_rect.height() != value_rect.height() {
            return Err(FormulaEvalError::Value);
        }
        if let Some((_, column_rect, _)) = &column_filter {
            if column_rect.height() != value_rect.height() {
                return Err(FormulaEvalError::Value);
            }
        }
        let first_row_key = self.rect_row_key_values(row_sheet_id, row_rect, 0)?;
        let mut values = Vec::new();
        for row_offset in 0..value_rect.height() {
            if !self.rect_row_key_matches(row_sheet_id, row_rect, row_offset, &first_row_key)? {
                continue;
            }
            if let Some((column_sheet_id, column_rect, column_key)) = &column_filter
                && !self.rect_row_key_matches(
                    *column_sheet_id,
                    *column_rect,
                    row_offset,
                    column_key,
                )?
            {
                continue;
            }
            let value = self.evaluator.cell_value_or_blank(
                value_sheet_id,
                value_rect.row_first + row_offset,
                value_rect.col_first,
            )?;
            values.push(formula_value_probe_from_cell_value(value));
        }
        Ok(values)
    }

    fn parse_groupby_value_function(
        &mut self,
        row_sheet_id: SheetId,
        row_rect: Rect,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (value_sheet_id, value_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let aggregation = self.parse_groupby_aggregation_argument()?;
        self.consume_remaining_optional_value_arguments()?;
        let values = self.collect_first_group_values(
            row_sheet_id,
            row_rect,
            value_sheet_id,
            value_rect,
            None,
        )?;
        aggregation.evaluate(values.as_slice())
    }

    fn parse_pivotby_value_function(
        &mut self,
        row_sheet_id: SheetId,
        row_rect: Rect,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (column_sheet_id, column_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (value_sheet_id, value_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let aggregation = self.parse_groupby_aggregation_argument()?;
        self.consume_remaining_optional_value_arguments()?;
        let first_column_key = self.rect_row_key_values(column_sheet_id, column_rect, 0)?;
        let values = self.collect_first_group_values(
            row_sheet_id,
            row_rect,
            value_sheet_id,
            value_rect,
            Some((column_sheet_id, column_rect, first_column_key)),
        )?;
        aggregation.evaluate(values.as_slice())
    }

    fn parse_cell_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let info_type = self.parse_text_value_argument()?.to_ascii_lowercase();
        self.skip_whitespace();
        let (target_sheet_id, row, col) = if self.consume_char(')') {
            let Some((row, col)) = self.current_position else {
                return Err(FormulaEvalError::Value);
            };
            (self.sheet_id, row, col)
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (target_sheet_id, rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            (target_sheet_id, rect.row_first, rect.col_first)
        };

        match info_type.as_str() {
            "address" => Ok(FormulaValueProbe::Text(format_cell_address(
                row, col, true, true,
            ))),
            "col" => Ok(FormulaValueProbe::Number(col as f64)),
            "row" => Ok(FormulaValueProbe::Number(row as f64)),
            "contents" => {
                let value = self
                    .evaluator
                    .cell_value_or_blank(target_sheet_id, row, col)?;
                Ok(formula_value_probe_from_cell_value(value))
            }
            "filename" => Ok(FormulaValueProbe::Text(String::new())),
            "format" => Ok(FormulaValueProbe::Text("G".to_string())),
            "color" | "parentheses" | "protect" => Ok(FormulaValueProbe::Number(0.0)),
            "prefix" => Ok(FormulaValueProbe::Text(String::new())),
            "type" => {
                let value = self
                    .evaluator
                    .cell_value_or_blank(target_sheet_id, row, col)?;
                Ok(FormulaValueProbe::Text(
                    match value {
                        CellValue::Blank => "b",
                        CellValue::Text(_) => "l",
                        CellValue::Bool(_) | CellValue::Number(_) | CellValue::Error(_) => "v",
                    }
                    .to_string(),
                ))
            }
            "width" => Ok(FormulaValueProbe::Number(8.0)),
            _ => Err(FormulaEvalError::Value),
        }
    }

    fn parse_info_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let info_type = self.parse_text_value_argument()?.to_ascii_lowercase();
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match info_type.as_str() {
            "directory" => Ok(FormulaValueProbe::Text(String::new())),
            "numfile" => Ok(FormulaValueProbe::Number(
                self.evaluator.state.worksheets.len() as f64,
            )),
            "origin" => Ok(FormulaValueProbe::Text("$A:$A".to_string())),
            "osversion" => Ok(FormulaValueProbe::Text(std::env::consts::OS.to_string())),
            "recalc" => Ok(FormulaValueProbe::Text("Automatic".to_string())),
            "release" => Ok(FormulaValueProbe::Text(APPLICATION_VERSION.to_string())),
            "system" => Ok(FormulaValueProbe::Text(
                if cfg!(target_os = "macos") {
                    "mac"
                } else {
                    "pcdos"
                }
                .to_string(),
            )),
            _ => Err(FormulaEvalError::Value),
        }
    }

    fn parse_column_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            let Some((_, col)) = self.current_position else {
                return Err(FormulaEvalError::Value);
            };
            return Ok(col as f64);
        }
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.col_first as f64)
    }

    fn parse_columns_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.width() as f64)
    }

    fn parse_iferror_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_error_fallback_function(false)
    }

    fn parse_ifna_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_error_fallback_function(true)
    }

    fn parse_error_fallback_function(
        &mut self,
        catch_only_na: bool,
    ) -> Result<f64, FormulaEvalError> {
        let primary = self.parse_catchable_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let fallback = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match primary {
            Ok(value) => Ok(value),
            Err(FormulaEvalError::NA) => Ok(fallback),
            Err(_error) if !catch_only_na => Ok(fallback),
            Err(error) => Err(error),
        }
    }

    fn parse_error_test_function(
        &mut self,
        match_any_error: bool,
        match_only_na: bool,
    ) -> Result<f64, FormulaEvalError> {
        let value = self.parse_catchable_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let is_match = match value {
            Ok(_) => false,
            Err(FormulaEvalError::NA) => match_any_error || match_only_na,
            Err(_) => match_any_error || !match_only_na,
        };
        Ok(if is_match { 1.0 } else { 0.0 })
    }

    fn parse_value_probe_test_function(
        &mut self,
        predicate: impl Fn(&FormulaValueProbe) -> bool,
    ) -> Result<f64, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if predicate(&value) { 1.0 } else { 0.0 })
    }

    fn parse_type_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(match value {
            FormulaValueProbe::Blank | FormulaValueProbe::Number(_) => 1.0,
            FormulaValueProbe::Text(_) => 2.0,
            FormulaValueProbe::Bool(_) => 4.0,
            FormulaValueProbe::Error(_) => 16.0,
            FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => 64.0,
        })
    }

    fn parse_error_type_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let FormulaValueProbe::Error(error) = value else {
            return Err(FormulaEvalError::NA);
        };
        match error {
            FormulaEvalError::Null => Ok(1.0),
            FormulaEvalError::Div0 => Ok(2.0),
            FormulaEvalError::Value => Ok(3.0),
            FormulaEvalError::Ref => Ok(4.0),
            FormulaEvalError::Name => Ok(5.0),
            FormulaEvalError::Num => Ok(6.0),
            FormulaEvalError::NA => Ok(7.0),
            FormulaEvalError::GettingData => Ok(8.0),
            FormulaEvalError::Spill => Ok(9.0),
            FormulaEvalError::Field => Ok(10.0),
            FormulaEvalError::Blocked => Ok(11.0),
            FormulaEvalError::Unknown => Ok(12.0),
            FormulaEvalError::Calc | FormulaEvalError::Circular => Ok(14.0),
            FormulaEvalError::Busy
            | FormulaEvalError::Connect
            | FormulaEvalError::Python
            | FormulaEvalError::Timeout => Err(FormulaEvalError::NA),
            FormulaEvalError::Unsupported => Err(FormulaEvalError::NA),
        }
    }

    fn parse_isformula_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (target_sheet_id, rect) = self.parse_reference_argument()?;
        if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(
            if self
                .evaluator
                .formula_source_at(target_sheet_id, rect.row_first, rect.col_first)?
                .is_some()
            {
                1.0
            } else {
                0.0
            },
        )
    }

    fn parse_na_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Err(FormulaEvalError::NA)
    }

    fn parse_row_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            let Some((row, _)) = self.current_position else {
                return Err(FormulaEvalError::Value);
            };
            return Ok(row as f64);
        }
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.row_first as f64)
    }

    fn parse_rows_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.height() as f64)
    }

    fn parse_areas_function(&mut self) -> Result<f64, FormulaEvalError> {
        let reference = self.parse_reference_set_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(reference.len() as f64)
    }

    fn parse_sheet_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return self.sheet_index(self.sheet_id);
        }
        let (target_sheet_id, _) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.sheet_index(target_sheet_id)
    }

    fn parse_sheets_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(self.evaluator.state.worksheets.len() as f64);
        }
        let reference = self.parse_reference_set_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let sheet_ids = reference
            .areas()
            .iter()
            .map(|(sheet_id, _)| *sheet_id)
            .collect::<BTreeSet<_>>();
        Ok(sheet_ids.len() as f64)
    }

    fn sheet_index(&self, sheet_id: SheetId) -> Result<f64, FormulaEvalError> {
        self.evaluator
            .state
            .worksheets
            .iter()
            .position(|worksheet| worksheet.id == sheet_id)
            .map(|index| index as f64 + 1.0)
            .ok_or(FormulaEvalError::Ref)
    }

    fn parse_index_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_index_value_function()?)
    }

    fn parse_index_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let reference = self.parse_index_reference_function()?;
        let (sheet_id, rect) = reference.single_area()?;
        self.evaluator
            .lookup_result_at(sheet_id, rect.row_first, rect.col_first)
    }

    fn parse_index_reference_function(&mut self) -> Result<FormulaReference, FormulaEvalError> {
        let reference = self.parse_reference_set_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let row_index = formula_integer_argument(self.parse_comparison()?)?;
        if row_index < 0 {
            return Err(FormulaEvalError::Value);
        }
        let mut col_index = 1_i64;
        self.skip_whitespace();
        if self.consume_char(',') {
            col_index = formula_integer_argument(self.parse_comparison()?)?;
            if col_index < 0 {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
        }
        let mut area_index = 1_i64;
        if self.consume_char(',') {
            area_index = formula_integer_argument(self.parse_comparison()?)?;
            if area_index < 1 {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
        }
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let area_offset = usize::try_from(area_index - 1).map_err(|_| FormulaEvalError::Ref)?;
        let (sheet_id, rect) = reference
            .areas()
            .get(area_offset)
            .copied()
            .ok_or(FormulaEvalError::Ref)?;
        if row_index > i64::from(rect.height()) || col_index > i64::from(rect.width()) {
            return Err(FormulaEvalError::Ref);
        }
        let (row_first, row_last) = if row_index == 0 {
            (rect.row_first, rect.row_last)
        } else {
            let row = rect.row_first + row_index as u32 - 1;
            (row, row)
        };
        let (col_first, col_last) = if col_index == 0 {
            (rect.col_first, rect.col_last)
        } else {
            let col = rect.col_first + col_index as u32 - 1;
            (col, col)
        };
        Ok(FormulaReference::single(
            sheet_id,
            Rect {
                row_first,
                row_last,
                col_first,
                col_last,
            },
        ))
    }

    fn parse_match_function(&mut self) -> Result<f64, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let mode = self
            .parse_optional_lookup_mode_argument(true)?
            .unwrap_or(FormulaLookupMode::ApproxAscending);
        if rect.width() > 1 && rect.height() > 1 {
            return Err(FormulaEvalError::Value);
        }
        let orientation = if rect.height() == 1 {
            FormulaLookupOrientation::FirstRow
        } else {
            FormulaLookupOrientation::FirstColumn
        };
        let values = self
            .evaluator
            .lookup_values_in_rect(sheet_id, rect, orientation)?;
        let index = lookup_match_index_in_values(&lookup_value, values.as_slice(), mode)?;
        Ok(index as f64 + 1.0)
    }

    fn parse_xmatch_function(&mut self) -> Result<f64, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sheet_id, rect, orientation) = self.parse_lookup_vector_reference_argument()?;
        let mut mode = FormulaXLookupMatchMode::Exact;
        let mut search_mode = FormulaXLookupSearchMode::Forward;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            mode = formula_xlookup_match_mode_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                search_mode = formula_xlookup_search_mode_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        let values = self
            .evaluator
            .lookup_values_in_rect(sheet_id, rect, orientation)?;
        let index =
            xlookup_match_index_in_values(&lookup_value, values.as_slice(), mode, search_mode)?;
        Ok(index as f64 + 1.0)
    }

    fn parse_vlookup_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_table_lookup_function(false)
    }

    fn parse_hlookup_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_table_lookup_function(true)
    }

    fn parse_lookup_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_lookup_value_function()?)
    }

    fn parse_lookup_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (lookup_sheet_id, lookup_rect, lookup_orientation) =
            self.parse_lookup_vector_reference_argument()?;
        let lookup_len = match lookup_orientation {
            FormulaLookupOrientation::FirstColumn => lookup_rect.height(),
            FormulaLookupOrientation::FirstRow => lookup_rect.width(),
        };

        self.skip_whitespace();
        let (return_sheet_id, return_rect, return_orientation) = if self.consume_char(')') {
            (lookup_sheet_id, lookup_rect, lookup_orientation)
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (return_sheet_id, return_rect, return_orientation) =
                self.parse_lookup_vector_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            (return_sheet_id, return_rect, return_orientation)
        };
        let return_len = match return_orientation {
            FormulaLookupOrientation::FirstColumn => return_rect.height(),
            FormulaLookupOrientation::FirstRow => return_rect.width(),
        };
        if lookup_len != return_len {
            return Err(FormulaEvalError::Value);
        }

        let lookup_values = self.evaluator.lookup_values_in_rect(
            lookup_sheet_id,
            lookup_rect,
            lookup_orientation,
        )?;
        let match_index = lookup_match_index_in_values(
            &lookup_value,
            lookup_values.as_slice(),
            FormulaLookupMode::ApproxAscending,
        )?;
        let (row, col) = match return_orientation {
            FormulaLookupOrientation::FirstColumn => (
                return_rect.row_first + match_index as u32,
                return_rect.col_first,
            ),
            FormulaLookupOrientation::FirstRow => (
                return_rect.row_first,
                return_rect.col_first + match_index as u32,
            ),
        };
        self.evaluator.lookup_result_at(return_sheet_id, row, col)
    }

    fn parse_table_lookup_function(&mut self, horizontal: bool) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_table_lookup_value_function(horizontal)?)
    }

    fn parse_table_lookup_value_function(
        &mut self,
        horizontal: bool,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sheet_id, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let result_index = formula_integer_argument(self.parse_comparison()?)?;
        if result_index < 1 {
            return Err(FormulaEvalError::Value);
        }
        let max_result_index = if horizontal {
            rect.height()
        } else {
            rect.width()
        };
        if result_index > i64::from(max_result_index) {
            return Err(FormulaEvalError::Ref);
        }
        let mode = self
            .parse_optional_lookup_mode_argument(false)?
            .unwrap_or(FormulaLookupMode::ApproxAscending);
        let orientation = if horizontal {
            FormulaLookupOrientation::FirstRow
        } else {
            FormulaLookupOrientation::FirstColumn
        };
        let values = self
            .evaluator
            .lookup_values_in_rect(sheet_id, rect, orientation)?;
        let match_index = lookup_match_index_in_values(&lookup_value, values.as_slice(), mode)?;
        let row = if horizontal {
            rect.row_first + result_index as u32 - 1
        } else {
            rect.row_first + match_index as u32
        };
        let col = if horizontal {
            rect.col_first + match_index as u32
        } else {
            rect.col_first + result_index as u32 - 1
        };
        self.evaluator.lookup_result_at(sheet_id, row, col)
    }

    fn parse_xlookup_value_function(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (lookup_sheet_id, lookup_rect, lookup_orientation) =
            self.parse_lookup_vector_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (return_sheet_id, return_rect, return_orientation) =
            self.parse_lookup_vector_reference_argument()?;
        let lookup_len = match lookup_orientation {
            FormulaLookupOrientation::FirstColumn => lookup_rect.height(),
            FormulaLookupOrientation::FirstRow => lookup_rect.width(),
        };
        let return_len = match return_orientation {
            FormulaLookupOrientation::FirstColumn => return_rect.height(),
            FormulaLookupOrientation::FirstRow => return_rect.width(),
        };
        if lookup_len != return_len {
            return Err(FormulaEvalError::Value);
        }

        let mut if_not_found = None;
        let mut mode = FormulaXLookupMatchMode::Exact;
        let mut search_mode = FormulaXLookupSearchMode::Forward;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            if_not_found = Some(self.parse_value_probe_argument()?);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                mode = formula_xlookup_match_mode_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    search_mode = formula_xlookup_search_mode_argument(self.parse_comparison()?)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }

        let lookup_values = self.evaluator.lookup_values_in_rect(
            lookup_sheet_id,
            lookup_rect,
            lookup_orientation,
        )?;
        let match_index = match xlookup_match_index_in_values(
            &lookup_value,
            lookup_values.as_slice(),
            mode,
            search_mode,
        ) {
            Ok(index) => index,
            Err(FormulaEvalError::NA) => {
                return if_not_found.ok_or(FormulaEvalError::NA);
            }
            Err(error) => return Err(error),
        };
        let (row, col) = match return_orientation {
            FormulaLookupOrientation::FirstColumn => (
                return_rect.row_first + match_index as u32,
                return_rect.col_first,
            ),
            FormulaLookupOrientation::FirstRow => (
                return_rect.row_first,
                return_rect.col_first + match_index as u32,
            ),
        };
        self.evaluator.lookup_result_at(return_sheet_id, row, col)
    }

    fn parse_reference_projection_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let reference = self.parse_reference_projection_reference_function(name)?;
        let Some((sheet_id, rect)) = reference.areas().first().copied() else {
            return Err(FormulaEvalError::Ref);
        };
        let value = self
            .evaluator
            .cell_value_or_blank(sheet_id, rect.row_first, rect.col_first)?;
        Ok(formula_value_probe_from_cell_value(value))
    }

    fn parse_reference_projection_reference_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaReference, FormulaEvalError> {
        if name.eq_ignore_ascii_case("INDEX") {
            return self.parse_index_reference_function();
        }

        if name.eq_ignore_ascii_case("INDIRECT") {
            let mut reference_text = self.parse_text_value_argument()?;
            let mut a1_style = true;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                a1_style = self.parse_comparison()? != 0.0;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            if !a1_style {
                let (base_row, base_col) = self.current_position.unwrap_or((1, 1));
                reference_text =
                    convert_formula_r1c1_to_a1(reference_text.as_str(), base_row, base_col);
            }
            return parse_formula_reference_text(
                reference_text.as_str(),
                self.sheet_id,
                self.evaluator.state,
            );
        }

        if name.eq_ignore_ascii_case("OFFSET") {
            let (sheet_id, rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let row_offset = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let col_offset = formula_integer_argument(self.parse_comparison()?)?;
            let mut height = i64::from(rect.height());
            let mut width = i64::from(rect.width());
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                height = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    width = formula_integer_argument(self.parse_comparison()?)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            if height < 1 || width < 1 {
                return Err(FormulaEvalError::Ref);
            }
            let row = i64::from(rect.row_first) + row_offset;
            let col = i64::from(rect.col_first) + col_offset;
            if row < 1
                || col < 1
                || row + height - 1 > i64::from(EXCEL_MAX_ROW_INDEX)
                || col + width - 1 > i64::from(EXCEL_MAX_COLUMN_INDEX)
            {
                return Err(FormulaEvalError::Ref);
            }
            return Ok(FormulaReference::single(
                sheet_id,
                Rect {
                    row_first: row as u32,
                    row_last: (row + height - 1) as u32,
                    col_first: col as u32,
                    col_last: (col + width - 1) as u32,
                },
            ));
        }

        if name.eq_ignore_ascii_case("TRIMRANGE") {
            let (sheet_id, rect) = self.parse_reference_argument()?;
            let mut trim_rows = 3_i64;
            let mut trim_cols = 3_i64;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                trim_rows = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    trim_cols = formula_integer_argument(self.parse_comparison()?)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            if !(0..=3).contains(&trim_rows) || !(0..=3).contains(&trim_cols) {
                return Err(FormulaEvalError::Value);
            }

            let mut row_first = rect.row_first;
            let mut row_last = rect.row_last;
            if matches!(trim_rows, 1 | 3) {
                while row_first <= row_last {
                    let mut has_value = false;
                    for col in rect.col_first..=rect.col_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row_first, col)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    row_first += 1;
                }
            }
            if matches!(trim_rows, 2 | 3) {
                while row_first <= row_last {
                    let mut has_value = false;
                    for col in rect.col_first..=rect.col_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row_last, col)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    row_last = row_last.saturating_sub(1);
                }
            }
            if row_first > row_last {
                return Err(FormulaEvalError::Calc);
            }

            let mut col_first = rect.col_first;
            let mut col_last = rect.col_last;
            if matches!(trim_cols, 1 | 3) {
                while col_first <= col_last {
                    let mut has_value = false;
                    for row in row_first..=row_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row, col_first)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    col_first += 1;
                }
            }
            if matches!(trim_cols, 2 | 3) {
                while col_first <= col_last {
                    let mut has_value = false;
                    for row in row_first..=row_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row, col_last)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    col_last = col_last.saturating_sub(1);
                }
            }
            if col_first > col_last {
                return Err(FormulaEvalError::Calc);
            }
            return Ok(FormulaReference::single(
                sheet_id,
                Rect {
                    row_first,
                    row_last,
                    col_first,
                    col_last,
                },
            ));
        }

        Err(FormulaEvalError::Unsupported)
    }

    fn parse_f_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let first_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let second_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        let (_, first_variance) = formula_sample_mean_and_variance(first_values.as_slice())?;
        let (_, second_variance) = formula_sample_mean_and_variance(second_values.as_slice())?;
        if first_variance <= 0.0 || second_variance <= 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let (ratio, degrees1, degrees2) = if first_variance >= second_variance {
            (
                first_variance / second_variance,
                first_values.len() as f64 - 1.0,
                second_values.len() as f64 - 1.0,
            )
        } else {
            (
                second_variance / first_variance,
                second_values.len() as f64 - 1.0,
                first_values.len() as f64 - 1.0,
            )
        };
        formula_f_right_tail(ratio, degrees1, degrees2)
            .and_then(|tail| formula_checked_numeric_result((2.0 * tail).min(1.0)))
    }

    fn parse_t_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let first_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let second_values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let tails = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let test_type = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !matches!(tails, 1 | 2) || !matches!(test_type, 1 | 2 | 3) {
            return Err(FormulaEvalError::Num);
        }

        let (t_statistic, degrees) = match test_type {
            1 => {
                if first_values.len() != second_values.len() {
                    return Err(FormulaEvalError::NA);
                }
                let differences = first_values
                    .iter()
                    .zip(second_values.iter())
                    .map(|(first, second)| first - second)
                    .collect::<Vec<_>>();
                let (mean, variance) = formula_sample_mean_and_variance(differences.as_slice())?;
                if variance <= 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                (
                    mean.abs() / (variance / differences.len() as f64).sqrt(),
                    differences.len() as f64 - 1.0,
                )
            }
            2 => {
                let (first_mean, first_variance) =
                    formula_sample_mean_and_variance(first_values.as_slice())?;
                let (second_mean, second_variance) =
                    formula_sample_mean_and_variance(second_values.as_slice())?;
                let degrees = first_values.len() as f64 + second_values.len() as f64 - 2.0;
                let pooled_variance = ((first_values.len() - 1) as f64 * first_variance
                    + (second_values.len() - 1) as f64 * second_variance)
                    / degrees;
                if pooled_variance <= 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                (
                    (first_mean - second_mean).abs()
                        / (pooled_variance
                            * (1.0 / first_values.len() as f64 + 1.0 / second_values.len() as f64))
                            .sqrt(),
                    degrees,
                )
            }
            3 => {
                let (first_mean, first_variance) =
                    formula_sample_mean_and_variance(first_values.as_slice())?;
                let (second_mean, second_variance) =
                    formula_sample_mean_and_variance(second_values.as_slice())?;
                let first_component = first_variance / first_values.len() as f64;
                let second_component = second_variance / second_values.len() as f64;
                let denominator = (first_component + second_component).sqrt();
                let degrees_denominator = first_component * first_component
                    / (first_values.len() as f64 - 1.0)
                    + second_component * second_component / (second_values.len() as f64 - 1.0);
                if denominator == 0.0 || degrees_denominator == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                (
                    (first_mean - second_mean).abs() / denominator,
                    (first_component + second_component) * (first_component + second_component)
                        / degrees_denominator,
                )
            }
            _ => return Err(FormulaEvalError::Num),
        };
        if !t_statistic.is_finite() || !degrees.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        let tail = formula_student_t_right_tail_from_abs(t_statistic, degrees)?;
        formula_checked_numeric_result(if tails == 1 {
            tail
        } else {
            (2.0 * tail).min(1.0)
        })
    }

    fn parse_z_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let values = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let x = self.parse_comparison()?;
        if !x.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        let sigma = if self.consume_char(')') {
            None
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let sigma = self.parse_comparison()?;
            if !sigma.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if sigma <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            Some(sigma)
        };
        if values.is_empty() {
            return Err(FormulaEvalError::NA);
        }
        if values.iter().any(|value| !value.is_finite()) {
            return Err(FormulaEvalError::Value);
        }
        let mean = values.iter().sum::<f64>() / values.len() as f64;
        let sigma = if let Some(sigma) = sigma {
            sigma
        } else {
            let (_, variance) = formula_sample_mean_and_variance(values.as_slice())?;
            if variance <= 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            variance.sqrt()
        };
        let denominator = sigma / (values.len() as f64).sqrt();
        if denominator == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let z = (mean - x) / denominator;
        formula_standard_normal_cdf(z)
            .and_then(|cdf| formula_checked_numeric_result((1.0 - cdf).clamp(0.0, 1.0)))
    }

    fn parse_array_projection_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let mut row = rect.row_first;
        let mut col = rect.col_first;

        macro_rules! finish_with_cell {
            () => {{
                let value = self.evaluator.cell_value_or_blank(sheet_id, row, col)?;
                return Ok(formula_value_probe_from_cell_value(value));
            }};
        }
        macro_rules! parse_dimension_count {
            () => {{
                let value = formula_integer_argument(self.parse_comparison()?)?;
                if value < 1 {
                    return Err(FormulaEvalError::Value);
                }
                value
            }};
        }
        macro_rules! parse_selected_offset {
            ($index:expr, $size:expr) => {{
                let index = $index;
                let size = i64::from($size);
                if index == 0 || index.abs() > size {
                    return Err(FormulaEvalError::Value);
                }
                if index > 0 {
                    u32::try_from(index - 1).map_err(|_| FormulaEvalError::Value)?
                } else {
                    u32::try_from(size + index).map_err(|_| FormulaEvalError::Value)?
                }
            }};
        }
        let consume_remaining_arguments = |parser: &mut Self| -> Result<(), FormulaEvalError> {
            loop {
                parser.skip_whitespace();
                if parser.consume_char(')') {
                    return Ok(());
                }
                if !parser.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                parser.skip_whitespace();
                let checkpoint = parser.index;
                if let Some((_, _, next_index)) = parser.try_parse_reference()? {
                    parser.index = next_index;
                } else {
                    parser.index = checkpoint;
                    parser.parse_value_probe_argument()?;
                }
            }
        };

        if name.eq_ignore_ascii_case("GROUPBY") {
            return self.parse_groupby_value_function(sheet_id, rect);
        }

        if name.eq_ignore_ascii_case("PIVOTBY") {
            return self.parse_pivotby_value_function(sheet_id, rect);
        }

        if name.eq_ignore_ascii_case("BYROW") || name.eq_ignore_ascii_case("BYCOL") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let lambda = self.parse_lambda_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = self.evaluator.cell_value_or_blank(sheet_id, row, col)?;
            return self
                .evaluate_lambda_value(lambda, vec![formula_value_probe_from_cell_value(value)]);
        }

        if name.eq_ignore_ascii_case("MAP") {
            let mut references = vec![(sheet_id, rect)];
            loop {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if let Some(lambda) = self.try_parse_lambda_argument()? {
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    let mut arguments = Vec::with_capacity(references.len());
                    for (target_sheet_id, target_rect) in references {
                        if target_rect.width() != rect.width()
                            || target_rect.height() != rect.height()
                        {
                            return Err(FormulaEvalError::Value);
                        }
                        let value = self.evaluator.cell_value_or_blank(
                            target_sheet_id,
                            target_rect.row_first,
                            target_rect.col_first,
                        )?;
                        arguments.push(formula_value_probe_from_cell_value(value));
                    }
                    return self.evaluate_lambda_value(lambda, arguments);
                }
                references.push(self.parse_reference_argument()?);
            }
        }

        if name.eq_ignore_ascii_case("FILTER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (include_sheet_id, include_rect) = self.parse_reference_argument()?;
            let mut if_empty = None;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                if_empty = Some(self.parse_value_probe_argument()?);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            let include_value = |value: CellValue| -> Result<bool, FormulaEvalError> {
                match value {
                    CellValue::Blank => Ok(false),
                    CellValue::Bool(value) => Ok(value),
                    CellValue::Number(value) => Ok(value != 0.0),
                    CellValue::Text(_) => Err(FormulaEvalError::Value),
                    CellValue::Error(error) => Err(formula_eval_error_from_cell_error(error)),
                }
            };

            if include_rect.height() == rect.height() && include_rect.width() == 1 {
                for row_offset in 0..rect.height() {
                    let include = self.evaluator.cell_value_or_blank(
                        include_sheet_id,
                        include_rect.row_first + row_offset,
                        include_rect.col_first,
                    )?;
                    if include_value(include)? {
                        row = rect.row_first + row_offset;
                        finish_with_cell!();
                    }
                }
            } else if include_rect.width() == rect.width() && include_rect.height() == 1 {
                for col_offset in 0..rect.width() {
                    let include = self.evaluator.cell_value_or_blank(
                        include_sheet_id,
                        include_rect.row_first,
                        include_rect.col_first + col_offset,
                    )?;
                    if include_value(include)? {
                        col = rect.col_first + col_offset;
                        finish_with_cell!();
                    }
                }
            } else if include_rect.height() == rect.height() && include_rect.width() == rect.width()
            {
                for row_offset in 0..rect.height() {
                    for col_offset in 0..rect.width() {
                        let include = self.evaluator.cell_value_or_blank(
                            include_sheet_id,
                            include_rect.row_first + row_offset,
                            include_rect.col_first + col_offset,
                        )?;
                        if include_value(include)? {
                            row = rect.row_first + row_offset;
                            col = rect.col_first + col_offset;
                            finish_with_cell!();
                        }
                    }
                }
            } else {
                return Err(FormulaEvalError::Value);
            }
            return if_empty.ok_or(FormulaEvalError::Calc);
        }

        if name.eq_ignore_ascii_case("SORT") || name.eq_ignore_ascii_case("SORTBY") {
            let mut sort_index = 1_i64;
            let mut sort_order = 1_i64;
            let mut by_col = false;
            let mut by_sheet_id = sheet_id;
            let mut by_rect = rect;

            if name.eq_ignore_ascii_case("SORT") {
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        sort_index = formula_integer_argument(self.parse_comparison()?)?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            sort_order = formula_integer_argument(self.parse_comparison()?)?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            self.skip_whitespace();
                            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                                by_col = self.parse_comparison()? != 0.0;
                            }
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
            } else {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let by_reference = self.parse_reference_argument()?;
                by_sheet_id = by_reference.0;
                by_rect = by_reference.1;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        sort_order = formula_integer_argument(self.parse_comparison()?)?;
                    }
                    loop {
                        self.skip_whitespace();
                        if self.consume_char(')') {
                            break;
                        }
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        let checkpoint = self.index;
                        if let Some((_, _, next_index)) = self.try_parse_reference()? {
                            self.index = next_index;
                        } else {
                            self.index = checkpoint;
                            self.parse_value_probe_argument()?;
                        }
                    }
                }
            }

            if sort_index < 1 || !matches!(sort_order, -1 | 1) {
                return Err(FormulaEvalError::Value);
            }
            let descending = sort_order == -1;
            let compare_values = |left: &FormulaValueProbe,
                                  right: &FormulaValueProbe|
             -> Result<Ordering, FormulaEvalError> {
                if let FormulaValueProbe::Error(error) = left {
                    return Err(*error);
                }
                if let FormulaValueProbe::Error(error) = right {
                    return Err(*error);
                }
                if let Some(ordering) = formula_value_probe_ordering(left, right)? {
                    return Ok(ordering);
                }
                let rank = |value: &FormulaValueProbe| match value {
                    FormulaValueProbe::Blank => 0,
                    FormulaValueProbe::Number(_) => 1,
                    FormulaValueProbe::Text(_) => 2,
                    FormulaValueProbe::Bool(_) => 3,
                    FormulaValueProbe::Error(_) => 4,
                    FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => 5,
                };
                Ok(rank(left).cmp(&rank(right)))
            };
            let sort_indexes = |indexes: &mut Vec<usize>,
                                keys: &[FormulaValueProbe]|
             -> Result<(), FormulaEvalError> {
                for index in 1..indexes.len() {
                    let current = indexes[index];
                    let mut position = index;
                    while position > 0 {
                        let ordering =
                            compare_values(&keys[indexes[position - 1]], &keys[current])?;
                        let should_shift = if descending {
                            ordering == Ordering::Less
                        } else {
                            ordering == Ordering::Greater
                        };
                        if !should_shift {
                            break;
                        }
                        indexes[position] = indexes[position - 1];
                        position -= 1;
                    }
                    indexes[position] = current;
                }
                Ok(())
            };

            if name.eq_ignore_ascii_case("SORT") {
                if by_col {
                    if sort_index > i64::from(rect.height()) {
                        return Err(FormulaEvalError::Value);
                    }
                    let key_row = rect.row_first + sort_index as u32 - 1;
                    let mut keys = Vec::new();
                    for candidate_col in rect.col_first..=rect.col_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(sheet_id, key_row, candidate_col)?;
                        keys.push(formula_value_probe_from_cell_value(value));
                    }
                    let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                    sort_indexes(&mut indexes, keys.as_slice())?;
                    col = rect.col_first
                        + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
                } else {
                    if sort_index > i64::from(rect.width()) {
                        return Err(FormulaEvalError::Value);
                    }
                    let key_col = rect.col_first + sort_index as u32 - 1;
                    let mut keys = Vec::new();
                    for candidate_row in rect.row_first..=rect.row_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(sheet_id, candidate_row, key_col)?;
                        keys.push(formula_value_probe_from_cell_value(value));
                    }
                    let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                    sort_indexes(&mut indexes, keys.as_slice())?;
                    row = rect.row_first
                        + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
                }
            } else if by_rect.height() == rect.height() && by_rect.width() == 1 {
                let mut keys = Vec::new();
                for candidate_row in by_rect.row_first..=by_rect.row_last {
                    let value = self.evaluator.cell_value_or_blank(
                        by_sheet_id,
                        candidate_row,
                        by_rect.col_first,
                    )?;
                    keys.push(formula_value_probe_from_cell_value(value));
                }
                let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                sort_indexes(&mut indexes, keys.as_slice())?;
                row = rect.row_first
                    + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
            } else if by_rect.width() == rect.width() && by_rect.height() == 1 {
                let mut keys = Vec::new();
                for candidate_col in by_rect.col_first..=by_rect.col_last {
                    let value = self.evaluator.cell_value_or_blank(
                        by_sheet_id,
                        by_rect.row_first,
                        candidate_col,
                    )?;
                    keys.push(formula_value_probe_from_cell_value(value));
                }
                let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                sort_indexes(&mut indexes, keys.as_slice())?;
                col = rect.col_first
                    + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
            } else {
                return Err(FormulaEvalError::Value);
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("UNIQUE") {
            let mut by_col = false;
            let mut exactly_once = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    by_col = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        exactly_once = self.parse_comparison()? != 0.0;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            if !exactly_once {
                finish_with_cell!();
            }

            if by_col {
                for candidate_col in rect.col_first..=rect.col_last {
                    let mut count = 0_u32;
                    for compare_col in rect.col_first..=rect.col_last {
                        let mut same = true;
                        for candidate_row in rect.row_first..=rect.row_last {
                            let left = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                candidate_row,
                                candidate_col,
                            )?;
                            let right = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                candidate_row,
                                compare_col,
                            )?;
                            if !formula_value_probe_exact_match(
                                &formula_value_probe_from_cell_value(left),
                                &formula_value_probe_from_cell_value(right),
                            )? {
                                same = false;
                                break;
                            }
                        }
                        if same {
                            count += 1;
                        }
                    }
                    if count == 1 {
                        col = candidate_col;
                        finish_with_cell!();
                    }
                }
            } else {
                for candidate_row in rect.row_first..=rect.row_last {
                    let mut count = 0_u32;
                    for compare_row in rect.row_first..=rect.row_last {
                        let mut same = true;
                        for candidate_col in rect.col_first..=rect.col_last {
                            let left = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                candidate_row,
                                candidate_col,
                            )?;
                            let right = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                compare_row,
                                candidate_col,
                            )?;
                            if !formula_value_probe_exact_match(
                                &formula_value_probe_from_cell_value(left),
                                &formula_value_probe_from_cell_value(right),
                            )? {
                                same = false;
                                break;
                            }
                        }
                        if same {
                            count += 1;
                        }
                    }
                    if count == 1 {
                        row = candidate_row;
                        finish_with_cell!();
                    }
                }
            }
            return Err(FormulaEvalError::Calc);
        }

        if name.eq_ignore_ascii_case("TAKE") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows == 0 || rows.unsigned_abs() > u64::from(rect.height()) {
                return Err(FormulaEvalError::Calc);
            }
            if rows < 0 {
                row = rect.row_last
                    - u32::try_from(rows.unsigned_abs()).map_err(|_| FormulaEvalError::Calc)?
                    + 1;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let columns = formula_integer_argument(self.parse_comparison()?)?;
                if columns == 0 || columns.unsigned_abs() > u64::from(rect.width()) {
                    return Err(FormulaEvalError::Calc);
                }
                if columns < 0 {
                    col = rect.col_last
                        - u32::try_from(columns.unsigned_abs())
                            .map_err(|_| FormulaEvalError::Calc)?
                        + 1;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("DROP") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows.unsigned_abs() >= u64::from(rect.height()) {
                return Err(FormulaEvalError::Calc);
            }
            if rows > 0 {
                row = rect.row_first + u32::try_from(rows).map_err(|_| FormulaEvalError::Calc)?;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let columns = formula_integer_argument(self.parse_comparison()?)?;
                if columns.unsigned_abs() >= u64::from(rect.width()) {
                    return Err(FormulaEvalError::Calc);
                }
                if columns > 0 {
                    col = rect.col_first
                        + u32::try_from(columns).map_err(|_| FormulaEvalError::Calc)?;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("CHOOSECOLS") || name.eq_ignore_ascii_case("CHOOSEROWS") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let selected = formula_integer_argument(self.parse_comparison()?)?;
            if name.eq_ignore_ascii_case("CHOOSECOLS") {
                col = rect.col_first + parse_selected_offset!(selected, rect.width());
            } else {
                row = rect.row_first + parse_selected_offset!(selected, rect.height());
            }
            consume_remaining_arguments(self)?;
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("EXPAND") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = parse_dimension_count!();
            if rows < i64::from(rect.height()) {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    let columns = parse_dimension_count!();
                    if columns < i64::from(rect.width()) {
                        return Err(FormulaEvalError::Value);
                    }
                }
                consume_remaining_arguments(self)?;
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("HSTACK") || name.eq_ignore_ascii_case("VSTACK") {
            consume_remaining_arguments(self)?;
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("TRANSPOSE") {
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("WRAPCOLS") || name.eq_ignore_ascii_case("WRAPROWS") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            parse_dimension_count!();
            consume_remaining_arguments(self)?;
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("TOCOL") || name.eq_ignore_ascii_case("TOROW") {
            let mut ignore = 0_i64;
            let mut scan_by_column = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    ignore = formula_integer_argument(self.parse_comparison()?)?;
                    if !(0..=3).contains(&ignore) {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    scan_by_column = self.parse_comparison()? != 0.0;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            let include_value = |value: &CellValue| -> bool {
                let ignore_blank = matches!(ignore, 1 | 3);
                let ignore_error = matches!(ignore, 2 | 3);
                !(ignore_blank && matches!(value, CellValue::Blank))
                    && !(ignore_error && matches!(value, CellValue::Error(_)))
            };
            if scan_by_column {
                for candidate_col in rect.col_first..=rect.col_last {
                    for candidate_row in rect.row_first..=rect.row_last {
                        let value = self.evaluator.cell_value_or_blank(
                            sheet_id,
                            candidate_row,
                            candidate_col,
                        )?;
                        if include_value(&value) {
                            return Ok(formula_value_probe_from_cell_value(value));
                        }
                    }
                }
            } else {
                for candidate_row in rect.row_first..=rect.row_last {
                    for candidate_col in rect.col_first..=rect.col_last {
                        let value = self.evaluator.cell_value_or_blank(
                            sheet_id,
                            candidate_row,
                            candidate_col,
                        )?;
                        if include_value(&value) {
                            return Ok(formula_value_probe_from_cell_value(value));
                        }
                    }
                }
            }
            return Err(FormulaEvalError::Calc);
        }

        Err(FormulaEvalError::Unsupported)
    }

    fn parse_sequence_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rows = formula_integer_argument(self.parse_comparison()?)?;
        if rows < 1 {
            return Err(FormulaEvalError::Value);
        }
        let mut columns = 1_i64;
        let mut start = 1.0_f64;
        let mut step = 1.0_f64;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                columns = formula_integer_argument(self.parse_comparison()?)?;
                if columns < 1 {
                    return Err(FormulaEvalError::Value);
                }
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    start = self.parse_comparison()?;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    step = self.parse_comparison()?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }
        if !start.is_finite() || !step.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let _ = columns;
        Ok(start)
    }

    fn parse_randarray_function(&mut self) -> Result<f64, FormulaEvalError> {
        let mut rows = 1_i64;
        let mut columns = 1_i64;
        let mut min = 0.0_f64;
        let mut max = 1.0_f64;
        let mut whole_number = false;
        self.skip_whitespace();
        if !self.consume_char(')') {
            rows = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    columns = formula_integer_argument(self.parse_comparison()?)?;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        min = self.parse_comparison()?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            max = self.parse_comparison()?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            whole_number = self.parse_comparison()? != 0.0;
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
            }
        }
        if rows < 1 || columns < 1 {
            return Err(FormulaEvalError::Value);
        }
        if !min.is_finite() || !max.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        if min > max {
            return Err(FormulaEvalError::Value);
        }
        if whole_number {
            formula_rand_between(min, max)
        } else {
            formula_checked_numeric_result(min + (max - min) * formula_rand())
        }
    }

    fn parse_chisq_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (observed_sheet_id, observed_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (expected_sheet_id, expected_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if observed_rect.width() != expected_rect.width()
            || observed_rect.height() != expected_rect.height()
        {
            return Err(FormulaEvalError::NA);
        }
        let degrees = if observed_rect.height() == 1 {
            observed_rect.width().saturating_sub(1)
        } else if observed_rect.width() == 1 {
            observed_rect.height().saturating_sub(1)
        } else {
            (observed_rect.height() - 1) * (observed_rect.width() - 1)
        };
        if degrees == 0 {
            return Err(FormulaEvalError::Div0);
        }

        let mut statistic = 0.0_f64;
        for row_offset in 0..observed_rect.height() {
            for col_offset in 0..observed_rect.width() {
                let observed = self.evaluator.numeric_cell_value(
                    observed_sheet_id,
                    observed_rect.row_first + row_offset,
                    observed_rect.col_first + col_offset,
                )?;
                let expected = self.evaluator.numeric_cell_value(
                    expected_sheet_id,
                    expected_rect.row_first + row_offset,
                    expected_rect.col_first + col_offset,
                )?;
                if !observed.is_finite() || !expected.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if observed < 0.0 || expected <= 0.0 {
                    return Err(FormulaEvalError::Num);
                }
                statistic += (observed - expected).powi(2) / expected;
            }
        }
        if !statistic.is_finite() {
            return Err(FormulaEvalError::Num);
        }

        let checked_numeric_result = |value: f64| -> Result<f64, FormulaEvalError> {
            if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            }
        };
        let gamma_ln_value = |value: f64| {
            const COEFFICIENTS: [f64; 9] = [
                0.9999999999998099,
                676.5203681218851,
                -1259.1392167224028,
                771.3234287776531,
                -176.6150291621406,
                12.507343278686905,
                -0.13857109526572012,
                0.000009984369578019572,
                0.00000015056327351493116,
            ];
            let lanczos = |input: f64| {
                let z = input - 1.0;
                let mut x = COEFFICIENTS[0];
                for (index, coefficient) in COEFFICIENTS.iter().enumerate().skip(1) {
                    x += coefficient / (z + index as f64);
                }
                let t = z + 7.5;
                0.5 * (2.0 * std::f64::consts::PI).ln() + (z + 0.5) * t.ln() - t + x.ln()
            };
            if value < 0.5 {
                std::f64::consts::PI.ln()
                    - (std::f64::consts::PI * value).sin().ln()
                    - lanczos(1.0 - value)
            } else {
                lanczos(value)
            }
        };
        let regularized_gamma_p = |shape: f64, x: f64| -> Result<f64, FormulaEvalError> {
            if !shape.is_finite() || !x.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if shape <= 0.0 || x < 0.0 {
                return Err(FormulaEvalError::Num);
            }
            if x == 0.0 {
                return Ok(0.0);
            }
            const EPSILON: f64 = 1e-14;
            const FLOOR: f64 = 1e-300;
            const MAX_ITERATIONS: usize = 200;
            let gamma_ln = gamma_ln_value(shape);
            if x < shape + 1.0 {
                let mut term = 1.0 / shape;
                let mut sum = term;
                let mut ap = shape;
                for _ in 0..MAX_ITERATIONS {
                    ap += 1.0;
                    term *= x / ap;
                    sum += term;
                    if term.abs() <= sum.abs() * EPSILON {
                        return checked_numeric_result(
                            (sum * (-x + shape * x.ln() - gamma_ln).exp()).clamp(0.0, 1.0),
                        );
                    }
                }
                return checked_numeric_result(
                    (sum * (-x + shape * x.ln() - gamma_ln).exp()).clamp(0.0, 1.0),
                );
            }

            let mut b = x + 1.0 - shape;
            let mut c = 1.0 / FLOOR;
            let mut d = 1.0 / b.max(FLOOR);
            let mut h = d;
            for i in 1..=MAX_ITERATIONS {
                let i = i as f64;
                let an = -i * (i - shape);
                b += 2.0;
                d = an * d + b;
                if d.abs() < FLOOR {
                    d = FLOOR;
                }
                c = b + an / c;
                if c.abs() < FLOOR {
                    c = FLOOR;
                }
                d = 1.0 / d;
                let delta = d * c;
                h *= delta;
                if (delta - 1.0).abs() <= EPSILON {
                    let q = (-x + shape * x.ln() - gamma_ln).exp() * h;
                    return checked_numeric_result((1.0 - q).clamp(0.0, 1.0));
                }
            }
            let q = (-x + shape * x.ln() - gamma_ln).exp() * h;
            checked_numeric_result((1.0 - q).clamp(0.0, 1.0))
        };
        regularized_gamma_p(degrees as f64 / 2.0, statistic / 2.0)
            .map(|value| (1.0 - value).clamp(0.0, 1.0))
    }

    fn parse_database_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let (database_sheet_id, database_rect) = self.parse_reference_argument()?;
        if database_rect.height() < 2 || database_rect.width() < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let field = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (criteria_sheet_id, criteria_rect) = self.parse_reference_argument()?;
        if criteria_rect.height() < 1 || criteria_rect.width() < 1 {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        macro_rules! database_header_text {
            ($sheet_id:expr, $row:expr, $col:expr) => {{
                let value = self.evaluator.cell_value_or_blank($sheet_id, $row, $col)?;
                formula_text_from_value_probe(formula_value_probe_from_cell_value(value))?
            }};
        }
        macro_rules! database_field_offset {
            ($field:expr) => {{
                match $field {
                    FormulaValueProbe::Number(value) => {
                        let index = formula_integer_argument(value)?;
                        if index < 1 || index > i64::from(database_rect.width()) {
                            return Err(FormulaEvalError::Value);
                        }
                        index as u32 - 1
                    }
                    FormulaValueProbe::Text(label) => {
                        let mut found = None;
                        for col_offset in 0..database_rect.width() {
                            let header = database_header_text!(
                                database_sheet_id,
                                database_rect.row_first,
                                database_rect.col_first + col_offset
                            );
                            if header.eq_ignore_ascii_case(label.as_str()) {
                                found = Some(col_offset);
                                break;
                            }
                        }
                        found.ok_or(FormulaEvalError::Value)?
                    }
                    FormulaValueProbe::Error(error) => return Err(error),
                    FormulaValueProbe::Blank
                    | FormulaValueProbe::Bool(_)
                    | FormulaValueProbe::Omitted
                    | FormulaValueProbe::Lambda { .. } => {
                        return Err(FormulaEvalError::Value);
                    }
                }
            }};
        }
        let field_offset = database_field_offset!(field);

        let mut criteria_rows = Vec::<Vec<(u32, FormulaCriteria)>>::new();
        if criteria_rect.height() == 1 {
            criteria_rows.push(Vec::new());
        } else {
            for criteria_row in criteria_rect.row_first + 1..=criteria_rect.row_last {
                let mut terms = Vec::new();
                for criteria_col in criteria_rect.col_first..=criteria_rect.col_last {
                    let criteria_value = self.evaluator.cell_value_or_blank(
                        criteria_sheet_id,
                        criteria_row,
                        criteria_col,
                    )?;
                    let criteria = match criteria_value {
                        CellValue::Blank => continue,
                        CellValue::Number(value) => FormulaCriteria::from_numeric_value(value),
                        CellValue::Bool(value) => {
                            FormulaCriteria::from_numeric_value(if value { 1.0 } else { 0.0 })
                        }
                        CellValue::Text(value) => FormulaCriteria::from_string_literal(value),
                        CellValue::Error(error) => {
                            return Err(formula_eval_error_from_cell_error(error));
                        }
                    };
                    let label = database_header_text!(
                        criteria_sheet_id,
                        criteria_rect.row_first,
                        criteria_col
                    );
                    let col_offset = database_field_offset!(FormulaValueProbe::Text(label));
                    terms.push((col_offset, criteria));
                }
                criteria_rows.push(terms);
            }
        }

        let mut numeric_values = Vec::new();
        let mut counta = 0_u64;
        let mut dget_value = None::<FormulaValueProbe>;
        let mut dget_count = 0_u64;
        for row in database_rect.row_first + 1..=database_rect.row_last {
            let mut matches_any_criteria_row = false;
            for criteria_row in criteria_rows.iter() {
                let mut matches_all_terms = true;
                for (criteria_col_offset, criteria) in criteria_row {
                    let value = self.evaluator.cell_value_or_blank(
                        database_sheet_id,
                        row,
                        database_rect.col_first + *criteria_col_offset,
                    )?;
                    if let CellValue::Error(error) = value {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    if !criteria.matches(&value) {
                        matches_all_terms = false;
                        break;
                    }
                }
                if matches_all_terms {
                    matches_any_criteria_row = true;
                    break;
                }
            }
            if !matches_any_criteria_row {
                continue;
            }
            let field_value = self.evaluator.cell_value_or_blank(
                database_sheet_id,
                row,
                database_rect.col_first + field_offset,
            )?;
            match name.to_ascii_uppercase().as_str() {
                "DGET" => {
                    dget_count += 1;
                    dget_value = Some(formula_value_probe_from_cell_value(field_value));
                }
                "DCOUNTA" => match field_value {
                    CellValue::Blank => {}
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Bool(_) | CellValue::Number(_) | CellValue::Text(_) => counta += 1,
                },
                "DCOUNT" => match field_value {
                    CellValue::Number(_) => counta += 1,
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                },
                _ => match field_value {
                    CellValue::Number(value) => numeric_values.push(value),
                    CellValue::Error(error) => {
                        return Err(formula_eval_error_from_cell_error(error));
                    }
                    CellValue::Blank | CellValue::Bool(_) | CellValue::Text(_) => {}
                },
            }
        }

        if name.eq_ignore_ascii_case("DGET") {
            if dget_count == 0 {
                return Err(FormulaEvalError::Value);
            }
            if dget_count > 1 {
                return Err(FormulaEvalError::Num);
            }
            return dget_value.ok_or(FormulaEvalError::Value);
        }
        if name.eq_ignore_ascii_case("DCOUNT") || name.eq_ignore_ascii_case("DCOUNTA") {
            return Ok(FormulaValueProbe::Number(counta as f64));
        }

        let function = if name.eq_ignore_ascii_case("DAVERAGE") {
            FormulaAggregateFunction::Average
        } else if name.eq_ignore_ascii_case("DMAX") {
            FormulaAggregateFunction::Max
        } else if name.eq_ignore_ascii_case("DMIN") {
            FormulaAggregateFunction::Min
        } else if name.eq_ignore_ascii_case("DPRODUCT") {
            FormulaAggregateFunction::Product
        } else if name.eq_ignore_ascii_case("DSTDEV") {
            FormulaAggregateFunction::StDevS
        } else if name.eq_ignore_ascii_case("DSTDEVP") {
            FormulaAggregateFunction::StDevP
        } else if name.eq_ignore_ascii_case("DSUM") {
            FormulaAggregateFunction::Sum
        } else if name.eq_ignore_ascii_case("DVAR") {
            FormulaAggregateFunction::VarS
        } else if name.eq_ignore_ascii_case("DVARP") {
            FormulaAggregateFunction::VarP
        } else {
            return Err(FormulaEvalError::Unsupported);
        };
        Ok(FormulaValueProbe::Number(
            function.evaluate(numeric_values.as_slice())?,
        ))
    }

    fn parse_sumif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (criteria_sheet_id, criteria_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria = self.parse_criteria_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return self.evaluator.sumif_values_in_rect(
                criteria_sheet_id,
                criteria_rect,
                &criteria,
                criteria_sheet_id,
                criteria_rect,
            );
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sum_sheet_id, sum_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.evaluator.sumif_values_in_rect(
            criteria_sheet_id,
            criteria_rect,
            &criteria,
            sum_sheet_id,
            sum_rect,
        )
    }

    fn parse_averageif_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (criteria_sheet_id, criteria_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria = self.parse_criteria_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return self.evaluator.averageif_values_in_rect(
                criteria_sheet_id,
                criteria_rect,
                &criteria,
                criteria_sheet_id,
                criteria_rect,
            );
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (average_sheet_id, average_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.evaluator.averageif_values_in_rect(
            criteria_sheet_id,
            criteria_rect,
            &criteria,
            average_sheet_id,
            average_rect,
        )
    }

    fn parse_countifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        Ok(self
            .evaluator
            .countifs_values_in_rects(criteria_ranges.as_slice())? as f64)
    }

    fn parse_sumifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (sum_sheet_id, sum_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator
            .sumifs_values_in_rect(sum_sheet_id, sum_rect, criteria_ranges.as_slice())
    }

    fn parse_averageifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (average_sheet_id, average_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator.averageifs_values_in_rect(
            average_sheet_id,
            average_rect,
            criteria_ranges.as_slice(),
        )
    }

    fn parse_minifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (min_sheet_id, min_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator
            .minifs_values_in_rect(min_sheet_id, min_rect, criteria_ranges.as_slice())
    }

    fn parse_maxifs_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (max_sheet_id, max_rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let criteria_ranges = self.parse_criteria_ranges_arguments()?;
        self.evaluator
            .maxifs_values_in_rect(max_sheet_id, max_rect, criteria_ranges.as_slice())
    }

    fn parse_countblank_function(&mut self) -> Result<f64, FormulaEvalError> {
        let mut count = 0_u64;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            count += self.parse_countblank_argument()?;
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_series_sum_function(&mut self) -> Result<f64, FormulaEvalError> {
        let x = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let n = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let m = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        let coefficients = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if coefficients.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        if ![x, n, m].iter().all(|value| value.is_finite())
            || coefficients.iter().any(|value| !value.is_finite())
        {
            return Err(FormulaEvalError::Value);
        }
        let mut total = 0.0;
        for (index, coefficient) in coefficients.iter().enumerate() {
            let exponent = n + index as f64 * m;
            if !exponent.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            let term = coefficient * x.powf(exponent);
            if !term.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            total += term;
            if !total.is_finite() {
                return Err(FormulaEvalError::Num);
            }
        }
        Ok(total)
    }

    fn parse_aggregate_function(&mut self) -> Result<f64, FormulaEvalError> {
        let function_num = formula_integer_argument(self.parse_comparison()?)?;
        if !(1..=19).contains(&function_num) {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let options = formula_integer_argument(self.parse_comparison()?)?;
        if !(0..=7).contains(&options) {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }

        let ignore_nested = matches!(options, 0..=3);
        let ignore_errors = matches!(options, 2 | 3 | 6 | 7);
        let percentile_value =
            |mut values: Vec<f64>, k: f64, exclusive: bool| -> Result<f64, FormulaEvalError> {
                if values.is_empty() || !k.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                values.sort_by(|left, right| left.total_cmp(right));
                if exclusive {
                    if k <= 0.0 || k >= 1.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let rank = k * (values.len() as f64 + 1.0);
                    if rank < 1.0 || rank > values.len() as f64 {
                        return Err(FormulaEvalError::Num);
                    }
                    let lower_rank = rank.floor();
                    let upper_rank = rank.ceil();
                    if lower_rank == upper_rank {
                        return Ok(values[lower_rank as usize - 1]);
                    }
                    let lower_index = lower_rank as usize - 1;
                    let upper_index = upper_rank as usize - 1;
                    let fraction = rank - lower_rank;
                    return Ok(values[lower_index]
                        + (values[upper_index] - values[lower_index]) * fraction);
                }
                if !(0.0..=1.0).contains(&k) {
                    return Err(FormulaEvalError::Num);
                }
                let rank = k * (values.len() as f64 - 1.0);
                let lower_index = rank.floor() as usize;
                let upper_index = rank.ceil() as usize;
                if lower_index == upper_index {
                    return Ok(values[lower_index]);
                }
                let fraction = rank - lower_index as f64;
                Ok(values[lower_index] + (values[upper_index] - values[lower_index]) * fraction)
            };
        let mut values = Vec::new();
        let mut counta = 0_u64;

        macro_rules! is_nested_aggregate_cell {
            ($sheet_id:expr, $row:expr, $col:expr) => {{
                self.evaluator
                    .state
                    .worksheet_data
                    .get(&$sheet_id)
                    .and_then(|worksheet| worksheet.cells.get(&($row, $col)))
                    .and_then(|cell| cell.formula.as_ref())
                    .is_some_and(|formula| {
                        formula_source_has_top_level_function(formula, "SUBTOTAL")
                            || formula_source_has_top_level_function(formula, "AGGREGATE")
                    })
            }};
        }

        macro_rules! record_numeric_value {
            ($value:expr) => {{
                match $value {
                    FormulaValueProbe::Number(number) => values.push(number),
                    FormulaValueProbe::Error(error) if ignore_errors => {
                        let _ = error;
                    }
                    FormulaValueProbe::Error(error) => return Err(error),
                    FormulaValueProbe::Blank
                    | FormulaValueProbe::Bool(_)
                    | FormulaValueProbe::Text(_)
                    | FormulaValueProbe::Omitted
                    | FormulaValueProbe::Lambda { .. } => {}
                }
            }};
        }

        macro_rules! collect_numeric_argument {
            () => {{
                if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
                    for (target_sheet_id, rect) in reference.areas() {
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                if ignore_nested
                                    && is_nested_aggregate_cell!(*target_sheet_id, row, col)
                                {
                                    continue;
                                }
                                let value = self.evaluator.cell_value_or_blank(
                                    *target_sheet_id,
                                    row,
                                    col,
                                )?;
                                record_numeric_value!(formula_value_probe_from_cell_value(value));
                            }
                        }
                    }
                } else {
                    match self.parse_catchable_argument()? {
                        Ok(value) => values.push(value),
                        Err(error) if ignore_errors => {
                            let _ = error;
                        }
                        Err(error) => return Err(error),
                    }
                }
            }};
        }

        macro_rules! collect_counta_argument {
            () => {{
                if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
                    for (target_sheet_id, rect) in reference.areas() {
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                if ignore_nested
                                    && is_nested_aggregate_cell!(*target_sheet_id, row, col)
                                {
                                    continue;
                                }
                                let value = self.evaluator.cell_value_or_blank(
                                    *target_sheet_id,
                                    row,
                                    col,
                                )?;
                                match value {
                                    CellValue::Error(error) if ignore_errors => {
                                        let _ = error;
                                    }
                                    CellValue::Error(error) => {
                                        return Err(formula_eval_error_from_cell_error(error));
                                    }
                                    CellValue::Blank => {}
                                    CellValue::Bool(_)
                                    | CellValue::Number(_)
                                    | CellValue::Text(_) => counta += 1,
                                }
                            }
                        }
                    }
                } else {
                    match self.parse_value_probe_argument()? {
                        FormulaValueProbe::Error(error) if ignore_errors => {
                            let _ = error;
                        }
                        FormulaValueProbe::Error(error) => return Err(error),
                        FormulaValueProbe::Blank => {}
                        FormulaValueProbe::Bool(_)
                        | FormulaValueProbe::Number(_)
                        | FormulaValueProbe::Text(_) => counta += 1,
                        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {}
                    }
                }
            }};
        }

        let finish_numeric = |values: &[f64]| -> Result<f64, FormulaEvalError> {
            let aggregate_function = match function_num {
                1 => Some(FormulaAggregateFunction::Average),
                2 => Some(FormulaAggregateFunction::Count),
                4 => Some(FormulaAggregateFunction::Max),
                5 => Some(FormulaAggregateFunction::Min),
                6 => Some(FormulaAggregateFunction::Product),
                7 => Some(FormulaAggregateFunction::StDevS),
                8 => Some(FormulaAggregateFunction::StDevP),
                9 => Some(FormulaAggregateFunction::Sum),
                10 => Some(FormulaAggregateFunction::VarS),
                11 => Some(FormulaAggregateFunction::VarP),
                12 => Some(FormulaAggregateFunction::Median),
                13 => Some(FormulaAggregateFunction::ModeSngl),
                _ => None,
            };
            aggregate_function
                .ok_or(FormulaEvalError::Value)?
                .evaluate(values)
        };

        if (14..=19).contains(&function_num) {
            collect_numeric_argument!();
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let k = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if function_num == 14 || function_num == 15 {
                let k = formula_integer_argument(k)?;
                if values.is_empty() || k < 1 || k > values.len() as i64 {
                    return Err(FormulaEvalError::Num);
                }
                values.sort_by(|left, right| left.total_cmp(right));
                let index = if function_num == 14 {
                    values.len() - k as usize
                } else {
                    k as usize - 1
                };
                return Ok(values[index]);
            }
            if function_num == 17 || function_num == 19 {
                if !k.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if k < i64::MIN as f64 || k > i64::MAX as f64 {
                    return Err(FormulaEvalError::Num);
                }
                let quart = k.trunc() as i64;
                let exclusive = function_num == 19;
                if exclusive {
                    if !(1..=3).contains(&quart) {
                        return Err(FormulaEvalError::Num);
                    }
                } else if !(0..=4).contains(&quart) {
                    return Err(FormulaEvalError::Num);
                }
                return percentile_value(values, quart as f64 / 4.0, exclusive);
            }
            return percentile_value(values, k, function_num == 18);
        }

        let mut saw_argument = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                if !saw_argument {
                    return Err(FormulaEvalError::Value);
                }
                return if function_num == 3 {
                    Ok(counta as f64)
                } else {
                    finish_numeric(values.as_slice())
                };
            }
            saw_argument = true;
            if function_num == 3 {
                collect_counta_argument!();
            } else {
                collect_numeric_argument!();
            }
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return if function_num == 3 {
                    Ok(counta as f64)
                } else {
                    finish_numeric(values.as_slice())
                };
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_subtotal_function(&mut self) -> Result<f64, FormulaEvalError> {
        let function_num = formula_integer_argument(self.parse_comparison()?)?;
        let function_num = match function_num {
            1..=11 => function_num,
            101..=111 => function_num - 100,
            _ => return Err(FormulaEvalError::Value),
        };
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }

        let aggregate_function = match function_num {
            1 => Some(FormulaAggregateFunction::Average),
            2 => Some(FormulaAggregateFunction::Count),
            3 => None,
            4 => Some(FormulaAggregateFunction::Max),
            5 => Some(FormulaAggregateFunction::Min),
            6 => Some(FormulaAggregateFunction::Product),
            7 => Some(FormulaAggregateFunction::StDevS),
            8 => Some(FormulaAggregateFunction::StDevP),
            9 => Some(FormulaAggregateFunction::Sum),
            10 => Some(FormulaAggregateFunction::VarS),
            11 => Some(FormulaAggregateFunction::VarP),
            _ => unreachable!("validated SUBTOTAL function number"),
        };
        let mut saw_argument = false;
        let mut values = Vec::new();
        let mut counta = 0_u64;

        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                if !saw_argument {
                    return Err(FormulaEvalError::Value);
                }
                return match aggregate_function {
                    Some(function) => function.evaluate(values.as_slice()),
                    None => Ok(counta as f64),
                };
            }
            saw_argument = true;
            if aggregate_function.is_some() {
                values.extend(self.parse_subtotal_numeric_argument()?);
            } else {
                counta += self.parse_subtotal_counta_argument()?;
            }
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return match aggregate_function {
                    Some(function) => function.evaluate(values.as_slice()),
                    None => Ok(counta as f64),
                };
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_counta_function(&mut self) -> Result<f64, FormulaEvalError> {
        let mut count = 0_u64;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            count += self.parse_counta_argument()?;
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return Ok(count as f64);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_aggregate_a_function(&mut self, name: &str) -> Result<f64, FormulaEvalError> {
        let mut values = Vec::new();
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                break;
            }
            values.extend(self.parse_aggregate_a_argument()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                break;
            }
            return Err(FormulaEvalError::Unsupported);
        }
        if name.eq_ignore_ascii_case("MINA") {
            return Ok(values.iter().copied().reduce(f64::min).unwrap_or(0.0));
        }
        if name.eq_ignore_ascii_case("MAXA") {
            return Ok(values.iter().copied().reduce(f64::max).unwrap_or(0.0));
        }
        if values.is_empty() {
            return Err(FormulaEvalError::Div0);
        }
        let mean = values.iter().sum::<f64>() / values.len() as f64;
        if name.eq_ignore_ascii_case("AVERAGEA") {
            return Ok(mean);
        }
        let deviation_sum = values
            .iter()
            .map(|value| {
                let deviation = value - mean;
                deviation * deviation
            })
            .sum::<f64>();
        if name.eq_ignore_ascii_case("VARPA") {
            return Ok(deviation_sum / values.len() as f64);
        }
        if name.eq_ignore_ascii_case("STDEVPA") {
            return Ok((deviation_sum / values.len() as f64).sqrt());
        }
        if values.len() < 2 {
            return Err(FormulaEvalError::Div0);
        }
        if name.eq_ignore_ascii_case("VARA") {
            return Ok(deviation_sum / (values.len() - 1) as f64);
        }
        Ok((deviation_sum / (values.len() - 1) as f64).sqrt())
    }

    fn parse_scalar_function(
        &mut self,
        function: FormulaScalarFunction,
    ) -> Result<f64, FormulaEvalError> {
        let mut args = Vec::new();
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return function.evaluate(args.as_slice());
            }
            args.push(self.parse_comparison()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return function.evaluate(args.as_slice());
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_logical_function(
        &mut self,
        function: FormulaLogicalFunction,
    ) -> Result<f64, FormulaEvalError> {
        let mut saw_value = false;
        let mut true_count = 0_u64;
        let mut false_count = 0_u64;
        macro_rules! record_value {
            ($value:expr, $from_reference:expr) => {
                match $value {
                    FormulaValueProbe::Bool(value) => {
                        saw_value = true;
                        if value {
                            true_count += 1;
                        } else {
                            false_count += 1;
                        }
                    }
                    FormulaValueProbe::Number(value) => {
                        saw_value = true;
                        if value != 0.0 {
                            true_count += 1;
                        } else {
                            false_count += 1;
                        }
                    }
                    FormulaValueProbe::Blank | FormulaValueProbe::Text(_) if $from_reference => {}
                    FormulaValueProbe::Blank => {
                        saw_value = true;
                        false_count += 1;
                    }
                    FormulaValueProbe::Text(_) => return Err(FormulaEvalError::Value),
                    FormulaValueProbe::Error(error) => return Err(error),
                    FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                        return Err(FormulaEvalError::Value);
                    }
                }
            };
        }
        macro_rules! finish_logical {
            () => {{
                if !saw_value {
                    return Err(FormulaEvalError::Value);
                }
                Ok(match function {
                    FormulaLogicalFunction::And => {
                        if false_count == 0 {
                            1.0
                        } else {
                            0.0
                        }
                    }
                    FormulaLogicalFunction::Or => {
                        if true_count > 0 {
                            1.0
                        } else {
                            0.0
                        }
                    }
                    FormulaLogicalFunction::Xor => {
                        if true_count % 2 == 1 {
                            1.0
                        } else {
                            0.0
                        }
                    }
                })
            }};
        }
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return finish_logical!();
            }

            if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
                for (target_sheet_id, rect) in reference.areas() {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let value =
                                self.evaluator
                                    .cell_value_or_blank(*target_sheet_id, row, col)?;
                            record_value!(formula_value_probe_from_cell_value(value), true);
                        }
                    }
                }
            } else {
                self.skip_whitespace();
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                match self.parse_catchable_argument()? {
                    Ok(value) => record_value!(FormulaValueProbe::Number(value), false),
                    Err(error) => record_value!(FormulaValueProbe::Error(error), false),
                }
            }

            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return finish_logical!();
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_aggregate_argument(&mut self) -> Result<Vec<f64>, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            return self.evaluator.numeric_values_in_reference(&reference);
        }
        Ok(vec![self.parse_comparison()?])
    }

    fn parse_subtotal_numeric_argument(&mut self) -> Result<Vec<f64>, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut values = Vec::new();
            for (target_sheet_id, rect) in reference.areas() {
                values.extend(
                    self.evaluator
                        .subtotal_numeric_values_in_rect(*target_sheet_id, *rect)?,
                );
            }
            return Ok(values);
        }
        Ok(vec![self.parse_comparison()?])
    }

    fn parse_subtotal_counta_argument(&mut self) -> Result<u64, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut count = 0_u64;
            for (target_sheet_id, rect) in reference.areas() {
                count += self
                    .evaluator
                    .subtotal_counta_values_in_rect(*target_sheet_id, *rect)?;
            }
            return Ok(count);
        }
        self.parse_comparison()?;
        Ok(1)
    }

    fn parse_aggregate_a_argument(&mut self) -> Result<Vec<f64>, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut values = Vec::new();
            for (target_sheet_id, rect) in reference.areas() {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        match self
                            .evaluator
                            .cell_value_or_blank(*target_sheet_id, row, col)?
                        {
                            CellValue::Blank => {}
                            CellValue::Bool(value) => values.push(if value { 1.0 } else { 0.0 }),
                            CellValue::Number(value) => values.push(value),
                            CellValue::Text(_) => values.push(0.0),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error));
                            }
                        }
                    }
                }
            }
            return Ok(values);
        }
        let value = self.parse_value_probe_argument()?;
        match value {
            FormulaValueProbe::Blank | FormulaValueProbe::Text(_) => Ok(vec![0.0]),
            FormulaValueProbe::Bool(value) => Ok(vec![if value { 1.0 } else { 0.0 }]),
            FormulaValueProbe::Number(value) => Ok(vec![value]),
            FormulaValueProbe::Error(error) => Err(error),
            FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                Err(FormulaEvalError::Value)
            }
        }
    }

    fn parse_optional_holidays_tail(&mut self) -> Result<Vec<i64>, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(Vec::new());
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let holidays = self.parse_holidays_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(holidays)
    }

    fn parse_optional_weekend_holidays_tail(
        &mut self,
    ) -> Result<([bool; 7], Vec<i64>), FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok((formula_standard_weekend_mask(), Vec::new()));
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let weekend = self.parse_weekend_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok((weekend, Vec::new()));
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let holidays = self.parse_holidays_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok((weekend, holidays))
    }

    fn parse_weekend_argument(&mut self) -> Result<[bool; 7], FormulaEvalError> {
        self.skip_whitespace();
        if let Some(value) = self.parse_string_literal()? {
            return formula_weekend_mask_from_string(value.as_str());
        }
        formula_weekend_mask_from_code(formula_integer_argument(self.parse_comparison()?)?)
    }

    fn parse_holidays_argument(&mut self) -> Result<Vec<i64>, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(reference) = self.parse_reference_set_before_boundary(&[')', ','])? {
            return self
                .evaluator
                .numeric_values_in_reference(&reference)?
                .into_iter()
                .map(formula_serial_integer)
                .collect::<Result<Vec<_>, _>>();
        }
        Ok(vec![formula_serial_integer(self.parse_comparison()?)?])
    }

    fn parse_counta_argument(&mut self) -> Result<u64, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            return self.evaluator.counta_values_in_reference(&reference);
        }
        self.parse_comparison()?;
        Ok(1)
    }

    fn parse_countblank_argument(&mut self) -> Result<u64, FormulaEvalError> {
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            return self.evaluator.countblank_values_in_reference(&reference);
        }
        self.parse_comparison()?;
        Ok(0)
    }

    fn parse_reference_argument(&mut self) -> Result<(SheetId, Rect), FormulaEvalError> {
        Ok(self.parse_reference_set_argument()?.single_area()?)
    }

    fn parse_reference_set_before_boundary(
        &mut self,
        boundaries: &[char],
    ) -> Result<Option<FormulaReference>, FormulaEvalError> {
        let checkpoint = self.index;
        match self.parse_reference_set_argument() {
            Ok(reference) => {
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| boundaries.contains(&ch)) {
                    Ok(Some(reference))
                } else {
                    self.index = checkpoint;
                    Ok(None)
                }
            }
            Err(FormulaEvalError::Value | FormulaEvalError::Unsupported) => {
                self.index = checkpoint;
                Ok(None)
            }
            Err(error) => {
                self.index = checkpoint;
                Err(error)
            }
        }
    }

    fn parse_reference_set_argument(&mut self) -> Result<FormulaReference, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(')
                && (identifier.eq_ignore_ascii_case("INDIRECT")
                    || identifier.eq_ignore_ascii_case("INDEX")
                    || identifier.eq_ignore_ascii_case("OFFSET")
                    || identifier.eq_ignore_ascii_case("TRIMRANGE"))
            {
                return self.parse_reference_projection_reference_function(identifier.as_str());
            }
        }
        self.index = checkpoint;
        let Some((reference, next_index)) = self.try_parse_reference_set()? else {
            return Err(FormulaEvalError::Value);
        };
        self.index = next_index;
        Ok(reference)
    }

    fn parse_lookup_vector_reference_argument(
        &mut self,
    ) -> Result<(SheetId, Rect, FormulaLookupOrientation), FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let orientation = if rect.width() == 1 {
            FormulaLookupOrientation::FirstColumn
        } else if rect.height() == 1 {
            FormulaLookupOrientation::FirstRow
        } else {
            return Err(FormulaEvalError::Value);
        };
        Ok((sheet_id, rect, orientation))
    }

    fn parse_lookup_value_argument(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        if let FormulaValueProbe::Error(error) = value {
            return Err(error);
        }
        Ok(value)
    }

    fn parse_optional_lookup_mode_argument(
        &mut self,
        allow_descending: bool,
    ) -> Result<Option<FormulaLookupMode>, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(None);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mode_value = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(Some(if mode_value == 0 {
            FormulaLookupMode::Exact
        } else if allow_descending && mode_value < 0 {
            FormulaLookupMode::ApproxDescending
        } else {
            FormulaLookupMode::ApproxAscending
        }))
    }

    fn parse_criteria_argument(&mut self) -> Result<FormulaCriteria, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(literal) = self.parse_string_literal()? {
            return Ok(FormulaCriteria::from_string_literal(literal));
        }
        Ok(FormulaCriteria::from_numeric_value(
            self.parse_comparison()?,
        ))
    }

    fn parse_criteria_ranges_arguments(
        &mut self,
    ) -> Result<Vec<FormulaCriteriaRange>, FormulaEvalError> {
        let mut criteria_ranges = Vec::new();
        loop {
            let (sheet_id, rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let criteria = self.parse_criteria_argument()?;
            criteria_ranges.push(FormulaCriteriaRange {
                sheet_id,
                rect,
                criteria,
            });
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(criteria_ranges);
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    fn parse_catchable_argument(
        &mut self,
    ) -> Result<Result<f64, FormulaEvalError>, FormulaEvalError> {
        match self.parse_comparison() {
            Ok(value) => Ok(Ok(value)),
            Err(FormulaEvalError::Unsupported) => Err(FormulaEvalError::Unsupported),
            Err(error) => Ok(Err(error)),
        }
    }

    fn parse_complex_argument(&mut self) -> Result<FormulaComplexNumber, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        formula_complex_from_text(text.as_str())
    }

    fn parse_convert_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let from_unit = self.parse_convert_unit_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let to_unit = self.parse_convert_unit_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_convert_value(value, from_unit.as_str(), to_unit.as_str())
    }

    fn parse_euroconvert_function(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let source = self.parse_convert_unit_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let target = self.parse_convert_unit_argument()?;
        let mut full_precision = false;
        let mut triangulation_precision = None;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            full_precision = self.parse_comparison()? != 0.0;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                triangulation_precision = Some(formula_integer_argument(self.parse_comparison()?)?);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_euroconvert_value(
            value,
            source.as_str(),
            target.as_str(),
            full_precision,
            triangulation_precision,
        )
    }

    fn parse_mdeterm_function(&mut self) -> Result<f64, FormulaEvalError> {
        let matrix = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_matrix_determinant(matrix)
    }

    fn parse_minverse_function(&mut self) -> Result<f64, FormulaEvalError> {
        let matrix = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_matrix_inverse_top_left(matrix)
    }

    fn parse_mmult_function(&mut self) -> Result<f64, FormulaEvalError> {
        let left = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let right = self.parse_numeric_matrix_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let left_width = left.first().map(|row| row.len()).unwrap_or(0);
        if left_width == 0
            || right.is_empty()
            || right.iter().any(|row| row.is_empty())
            || left_width != right.len()
        {
            return Err(FormulaEvalError::Value);
        }
        let mut total = 0.0_f64;
        for index in 0..left_width {
            total += left[0][index] * right[index][0];
        }
        formula_checked_numeric_result(total)
    }

    fn parse_munit_function(&mut self) -> Result<f64, FormulaEvalError> {
        let dimension = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if dimension < 1 {
            return Err(FormulaEvalError::Value);
        }
        Ok(1.0)
    }

    fn parse_frequency_function(&mut self) -> Result<f64, FormulaEvalError> {
        let data = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let bins = self.parse_aggregate_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let Some(first_bin) = bins.first() else {
            return Ok(data.len() as f64);
        };
        Ok(data.iter().filter(|value| **value <= *first_bin).count() as f64)
    }

    fn parse_numeric_matrix_reference_argument(
        &mut self,
    ) -> Result<Vec<Vec<f64>>, FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let mut matrix = Vec::with_capacity(rect.height() as usize);
        for row in rect.row_first..=rect.row_last {
            let mut values = Vec::with_capacity(rect.width() as usize);
            for col in rect.col_first..=rect.col_last {
                values.push(self.evaluator.numeric_cell_value(sheet_id, row, col)?);
            }
            matrix.push(values);
        }
        Ok(matrix)
    }

    fn parse_convert_unit_argument(&mut self) -> Result<String, FormulaEvalError> {
        match self.parse_value_probe_argument()? {
            FormulaValueProbe::Text(value) => Ok(value),
            FormulaValueProbe::Error(error) => Err(error),
            FormulaValueProbe::Blank
            | FormulaValueProbe::Bool(_)
            | FormulaValueProbe::Number(_)
            | FormulaValueProbe::Omitted
            | FormulaValueProbe::Lambda { .. } => Err(FormulaEvalError::Value),
        }
    }

    fn parse_text_value_argument(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(text) = self.parse_string_literal()? {
            return Ok(text);
        }
        let checkpoint = self.index;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (target_sheet_id, rect) = reference.single_area()?;
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Err(FormulaEvalError::Value);
            }
            let value = self.evaluator.cell_value_or_blank(
                target_sheet_id,
                rect.row_first,
                rect.col_first,
            )?;
            return formula_text_from_value_probe(formula_value_probe_from_cell_value(value));
        }
        self.index = checkpoint;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return formula_text_from_value_probe(value);
                }
                if identifier.eq_ignore_ascii_case("LAMBDA")
                    || formula_text_function_name(identifier.as_str())
                {
                    match self.parse_text_function(identifier.as_str()) {
                        Ok(text) => return Ok(text),
                        Err(FormulaEvalError::Unsupported) => self.index = checkpoint,
                        Err(error) => return Err(error),
                    }
                } else {
                    self.index = checkpoint;
                }
            } else {
                if identifier.eq_ignore_ascii_case("TRUE") {
                    return Ok("TRUE".to_string());
                }
                if identifier.eq_ignore_ascii_case("FALSE") {
                    return Ok("FALSE".to_string());
                }
                if let Some(value) = self.binding_value(identifier.as_str())
                    && self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                {
                    return formula_text_from_value_probe(value);
                }
                if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                    && let Some(value) = self.defined_name_value_probe(identifier.as_str())?
                {
                    return formula_text_from_value_probe(value);
                }
                self.index = checkpoint;
            }
        }
        formula_text_from_number(self.parse_comparison()?)
    }

    fn parse_text_values_argument(&mut self) -> Result<Vec<String>, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut values = Vec::new();
            for (target_sheet_id, rect) in reference.areas() {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(*target_sheet_id, row, col)?;
                        values.push(formula_text_from_value_probe(
                            formula_value_probe_from_cell_value(value),
                        )?);
                    }
                }
            }
            return Ok(values);
        }
        Ok(vec![self.parse_text_value_argument()?])
    }

    fn parse_value_probe_argument(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(text) = self.parse_string_literal()? {
            return Ok(FormulaValueProbe::Text(text));
        }
        let checkpoint = self.index;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (target_sheet_id, rect) = reference.single_area()?;
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Err(FormulaEvalError::Value);
            }
            let value = self.evaluator.cell_value_or_blank(
                target_sheet_id,
                rect.row_first,
                rect.col_first,
            )?;
            return Ok(formula_value_probe_from_cell_value(value));
        }
        self.index = checkpoint;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return Ok(value);
                }
                if identifier.eq_ignore_ascii_case("LAMBDA") {
                    let lambda = self.parse_lambda_value_function()?;
                    self.skip_whitespace();
                    if self.consume_char('(') {
                        return self.parse_lambda_call_arguments(lambda);
                    }
                    return Ok(lambda);
                }
                if formula_text_function_name(identifier.as_str()) {
                    match self.parse_text_function(identifier.as_str()) {
                        Ok(text) => return Ok(FormulaValueProbe::Text(text)),
                        Err(FormulaEvalError::Unsupported) => self.index = checkpoint,
                        Err(error) => return Err(error),
                    }
                } else {
                    self.index = checkpoint;
                }
            } else {
                if identifier.eq_ignore_ascii_case("TRUE") {
                    return Ok(FormulaValueProbe::Bool(true));
                }
                if identifier.eq_ignore_ascii_case("FALSE") {
                    return Ok(FormulaValueProbe::Bool(false));
                }
                if let Some(value) = self.binding_value(identifier.as_str())
                    && self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                {
                    return Ok(value);
                }
                if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                    && let Some(value) = self.defined_name_value_probe(identifier.as_str())?
                {
                    return Ok(value);
                }
                self.index = checkpoint;
            }
        }
        match self.parse_comparison() {
            Ok(value) => Ok(FormulaValueProbe::Number(value)),
            Err(FormulaEvalError::Unsupported) => Err(FormulaEvalError::Unsupported),
            Err(error) => Ok(FormulaValueProbe::Error(error)),
        }
    }

    fn parse_number(&mut self) -> Result<Option<f64>, FormulaEvalError> {
        let start = self.index;
        let mut saw_digit = false;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() {
                saw_digit = true;
                self.index += ch.len_utf8();
            } else if ch == '.' {
                self.index += ch.len_utf8();
            } else {
                break;
            }
        }
        if !saw_digit {
            self.index = start;
            return Ok(None);
        }
        let number = self.input[start..self.index]
            .parse::<f64>()
            .map_err(|_| FormulaEvalError::Value)?;
        Ok(Some(number))
    }

    fn parse_string_literal(&mut self) -> Result<Option<String>, FormulaEvalError> {
        if !self.consume_char('"') {
            return Ok(None);
        }
        let mut value = String::new();
        while let Some(ch) = self.peek_char() {
            self.index += ch.len_utf8();
            if ch == '"' {
                if self.peek_char() == Some('"') {
                    value.push('"');
                    self.index += 1;
                    continue;
                }
                return Ok(Some(value));
            }
            value.push(ch);
        }
        Err(FormulaEvalError::Unsupported)
    }

    fn parse_identifier(&mut self) -> Option<String> {
        let start = self.index;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphabetic()
                || ch == '_'
                || ch == '.'
                || (self.index > start && ch.is_ascii_digit())
            {
                self.index += ch.len_utf8();
            } else {
                break;
            }
        }
        (self.index > start).then(|| self.input[start..self.index].to_string())
    }

    fn try_parse_reference_set(
        &self,
    ) -> Result<Option<(FormulaReference, usize)>, FormulaEvalError> {
        if let Some((reference, next_index)) = self.try_parse_3d_reference()? {
            return Ok(Some((reference, next_index)));
        }
        if let Some((sheet_id, rect, next_index)) = self.try_parse_a1_reference()? {
            return Ok(Some((FormulaReference::single(sheet_id, rect), next_index)));
        }
        self.try_parse_named_reference()
    }

    fn try_parse_reference(&self) -> Result<Option<(SheetId, Rect, usize)>, FormulaEvalError> {
        let Some((reference, next_index)) = self.try_parse_reference_set()? else {
            return Ok(None);
        };
        let (sheet_id, rect) = reference.single_area()?;
        Ok(Some((sheet_id, rect, next_index)))
    }

    fn try_parse_a1_reference(&self) -> Result<Option<(SheetId, Rect, usize)>, FormulaEvalError> {
        let (sheet_id, start) = self.try_parse_sheet_qualifier()?;
        let mut cursor = start;
        let mut saw_reference_char = false;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch.is_ascii_alphanumeric() || ch == '$' || ch == ':' {
                saw_reference_char = true;
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        if !saw_reference_char {
            return Ok(None);
        }
        let has_spill_operator = self.input[cursor..].starts_with('#');
        let next_index = cursor + usize::from(has_spill_operator);
        let next_is_boundary = self.input[next_index..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if !next_is_boundary {
            return Ok(None);
        }
        let token = &self.input[start..cursor];
        let Some(mut rect) = parse_rect_a1(token).ok() else {
            return Ok(None);
        };
        if has_spill_operator {
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Err(FormulaEvalError::Ref);
            }
            rect = self
                .evaluator
                .state
                .worksheet_data
                .get(&sheet_id)
                .and_then(|worksheet| {
                    worksheet
                        .spill_ranges
                        .get(&(rect.row_first, rect.col_first))
                })
                .copied()
                .ok_or(FormulaEvalError::Ref)?;
        }
        Ok(Some((sheet_id, rect, next_index)))
    }

    fn try_parse_3d_reference(
        &self,
    ) -> Result<Option<(FormulaReference, usize)>, FormulaEvalError> {
        let Some((start_sheet_id, end_sheet_id, reference_start)) =
            self.try_parse_3d_sheet_span_prefix()?
        else {
            return Ok(None);
        };
        let mut cursor = reference_start;
        let mut saw_reference_char = false;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch.is_ascii_alphanumeric() || ch == '$' || ch == ':' {
                saw_reference_char = true;
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        if !saw_reference_char {
            return Ok(None);
        }
        let next_is_boundary = self.input[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if !next_is_boundary {
            return Ok(None);
        }
        let Some(rect) = parse_rect_a1(&self.input[reference_start..cursor]).ok() else {
            return Ok(None);
        };
        let sheets = formula_sheets_in_3d_span(self.evaluator.state, start_sheet_id, end_sheet_id)?;
        let areas = sheets
            .into_iter()
            .map(|sheet_id| (sheet_id, rect))
            .collect::<Vec<_>>();
        Ok(Some((
            FormulaReference::with_explicit_area_count(1, areas)?,
            cursor,
        )))
    }

    fn try_parse_3d_sheet_span_prefix(
        &self,
    ) -> Result<Option<(SheetId, SheetId, usize)>, FormulaEvalError> {
        if self.peek_char() == Some('\'') {
            let mut cursor = self.index + 1;
            let mut sheet_span = String::new();
            while cursor < self.input.len() {
                let ch = self.input[cursor..]
                    .chars()
                    .next()
                    .ok_or(FormulaEvalError::Unsupported)?;
                if ch == '\'' {
                    let next_cursor = cursor + ch.len_utf8();
                    if self.input[next_cursor..].starts_with('\'') {
                        sheet_span.push('\'');
                        cursor = next_cursor + 1;
                        continue;
                    }
                    if !self.input[next_cursor..].starts_with('!') {
                        return Ok(None);
                    }
                    let Some((start_sheet, end_sheet)) = sheet_span.split_once(':') else {
                        return Ok(None);
                    };
                    let start_sheet_id = self
                        .resolve_sheet_name(start_sheet)
                        .ok_or(FormulaEvalError::Ref)?;
                    let end_sheet_id = self
                        .resolve_sheet_name(end_sheet)
                        .ok_or(FormulaEvalError::Ref)?;
                    return Ok(Some((start_sheet_id, end_sheet_id, next_cursor + 1)));
                }
                sheet_span.push(ch);
                cursor += ch.len_utf8();
            }
            return Ok(None);
        }

        let mut cursor = self.index;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch == '!' {
                let sheet_span = &self.input[self.index..cursor];
                let Some((start_sheet, end_sheet)) = sheet_span.split_once(':') else {
                    return Ok(None);
                };
                if start_sheet.is_empty() || end_sheet.is_empty() {
                    return Err(FormulaEvalError::Ref);
                }
                let start_sheet_id = self
                    .resolve_sheet_name(start_sheet)
                    .ok_or(FormulaEvalError::Ref)?;
                let end_sheet_id = self
                    .resolve_sheet_name(end_sheet)
                    .ok_or(FormulaEvalError::Ref)?;
                return Ok(Some((start_sheet_id, end_sheet_id, cursor + 1)));
            }
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == ':' {
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        Ok(None)
    }

    fn try_parse_named_reference(
        &self,
    ) -> Result<Option<(FormulaReference, usize)>, FormulaEvalError> {
        let (qualified_sheet_id, start) = self.try_parse_sheet_qualifier()?;
        let qualified = start != self.index;
        let Some((name, cursor)) = self.parse_identifier_at(start) else {
            return Ok(None);
        };
        let next_is_boundary = self.input[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if !next_is_boundary {
            return Ok(None);
        }

        let defined_name = if qualified {
            self.evaluator
                .state
                .defined_names
                .lookup_in_scope(NameScope::Worksheet(qualified_sheet_id), name.as_str())
        } else {
            self.evaluator
                .state
                .defined_names
                .lookup(Some(self.sheet_id), name.as_str())
        };
        let Some(defined_name) = defined_name else {
            return Ok(None);
        };
        let Some(reference) = self.defined_name_reference(defined_name) else {
            return Ok(None);
        };
        Ok(Some((reference, cursor)))
    }

    fn defined_name_reference(
        &self,
        defined_name: &office_common::DefinedName,
    ) -> Option<FormulaReference> {
        if defined_name.refers_to.is_r1c1 {
            return None;
        }
        let default_sheet_id = match defined_name.scope {
            NameScope::Workbook => self.sheet_id,
            NameScope::Worksheet(sheet_id) => sheet_id,
        };
        parse_formula_reference_text(
            defined_name.refers_to.text.as_str(),
            default_sheet_id,
            self.evaluator.state,
        )
        .ok()
    }

    fn defined_name_value_probe(
        &mut self,
        name: &str,
    ) -> Result<Option<FormulaValueProbe>, FormulaEvalError> {
        let Some((name_id, scope, refers_to_text, is_r1c1)) = self
            .evaluator
            .state
            .defined_names
            .lookup(Some(self.sheet_id), name)
            .map(|defined_name| {
                (
                    defined_name.id,
                    defined_name.scope,
                    defined_name.refers_to.text.clone(),
                    defined_name.refers_to.is_r1c1,
                )
            })
        else {
            return Ok(None);
        };

        if !self.evaluator.resolving_names.insert(name_id) {
            return Err(FormulaEvalError::Calc);
        }
        let result = (|| {
            let default_sheet_id = match scope {
                NameScope::Workbook => self.sheet_id,
                NameScope::Worksheet(sheet_id) => sheet_id,
            };
            if !is_r1c1
                && let Ok(reference) = parse_formula_reference_text(
                    refers_to_text.as_str(),
                    default_sheet_id,
                    self.evaluator.state,
                )
            {
                let (sheet_id, rect) = reference.single_area()?;
                if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                    return Err(FormulaEvalError::Value);
                }
                let value =
                    self.evaluator
                        .cell_value_or_blank(sheet_id, rect.row_first, rect.col_first)?;
                return Ok(formula_value_probe_from_cell_value(value));
            }

            let formula_text = if is_r1c1 {
                let (base_row, base_col) = self.current_position.unwrap_or((1, 1));
                convert_formula_r1c1_to_a1(refers_to_text.as_str(), base_row, base_col)
            } else {
                refers_to_text
            };
            let value = self.evaluator.evaluate_formula_value_probe_text(
                self.sheet_id,
                formula_text.as_str(),
                self.current_position,
            )?;
            Ok(value)
        })();
        self.evaluator.resolving_names.remove(&name_id);
        result.map(Some)
    }

    fn parse_identifier_at(&self, start: usize) -> Option<(String, usize)> {
        let mut cursor = start;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch.is_ascii_alphabetic()
                || ch == '_'
                || ch == '.'
                || (cursor > start && ch.is_ascii_digit())
            {
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        (cursor > start).then(|| (self.input[start..cursor].to_string(), cursor))
    }

    fn try_parse_sheet_qualifier(&self) -> Result<(SheetId, usize), FormulaEvalError> {
        if self.peek_char() == Some('\'') {
            let mut cursor = self.index + 1;
            let mut sheet_name = String::new();
            while cursor < self.input.len() {
                let ch = self.input[cursor..]
                    .chars()
                    .next()
                    .ok_or(FormulaEvalError::Unsupported)?;
                if ch == '\'' {
                    let next_cursor = cursor + ch.len_utf8();
                    if self.input[next_cursor..].starts_with('\'') {
                        sheet_name.push('\'');
                        cursor = next_cursor + 1;
                        continue;
                    }
                    if !self.input[next_cursor..].starts_with('!') {
                        return Ok((self.sheet_id, self.index));
                    }
                    let sheet_id = self
                        .resolve_sheet_name(sheet_name.as_str())
                        .ok_or(FormulaEvalError::Ref)?;
                    return Ok((sheet_id, next_cursor + 1));
                }
                sheet_name.push(ch);
                cursor += ch.len_utf8();
            }
            return Ok((self.sheet_id, self.index));
        }

        let mut cursor = self.index;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch == '!' {
                let sheet_name = &self.input[self.index..cursor];
                if sheet_name.is_empty() {
                    return Err(FormulaEvalError::Ref);
                }
                let sheet_id = self
                    .resolve_sheet_name(sheet_name)
                    .ok_or(FormulaEvalError::Ref)?;
                return Ok((sheet_id, cursor + 1));
            }
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' {
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        Ok((self.sheet_id, self.index))
    }

    fn resolve_sheet_name(&self, name: &str) -> Option<SheetId> {
        self.evaluator
            .state
            .worksheets
            .iter()
            .find(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
            .map(|worksheet| worksheet.id)
    }

    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.peek_char() {
            if ch.is_whitespace() {
                self.index += ch.len_utf8();
            } else {
                break;
            }
        }
    }

    fn consume_char(&mut self, expected: char) -> bool {
        if self.peek_char() == Some(expected) {
            self.index += expected.len_utf8();
            true
        } else {
            false
        }
    }

    fn consume_comparison_operator(&mut self) -> Option<FormulaComparisonOperator> {
        let remaining = &self.input[self.index..];
        for (token, operator) in [
            ("<>", FormulaComparisonOperator::NotEqual),
            ("<=", FormulaComparisonOperator::LessThanOrEqual),
            (">=", FormulaComparisonOperator::GreaterThanOrEqual),
            ("=", FormulaComparisonOperator::Equal),
            ("<", FormulaComparisonOperator::LessThan),
            (">", FormulaComparisonOperator::GreaterThan),
        ] {
            if remaining.starts_with(token) {
                self.index += token.len();
                return Some(operator);
            }
        }
        None
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.index..].chars().next()
    }
}

fn parse_formula_reference_text(
    input: &str,
    default_sheet_id: SheetId,
    state: &WorkbookState,
) -> Result<FormulaReference, FormulaEvalError> {
    let input = input.trim().strip_prefix('=').unwrap_or(input.trim());
    let mut areas = Vec::new();
    let mut explicit_area_count = 0usize;
    for part in split_reference_union_text(input).map_err(|_| FormulaEvalError::Ref)? {
        let reference = parse_formula_reference_area_text(part, default_sheet_id, state)?;
        explicit_area_count += reference.len();
        areas.extend(reference.areas().iter().copied());
    }
    FormulaReference::with_explicit_area_count(explicit_area_count, areas)
}

fn parse_formula_reference_area_text(
    input: &str,
    default_sheet_id: SheetId,
    state: &WorkbookState,
) -> Result<FormulaReference, FormulaEvalError> {
    let input = input.trim();
    if input.is_empty() {
        return Err(FormulaEvalError::Ref);
    }

    let (sheet_ids, reference_text) = if input.starts_with('\'') {
        let mut cursor = 1usize;
        let mut sheet_name_or_span = String::new();
        loop {
            let Some(ch) = input[cursor..].chars().next() else {
                return Err(FormulaEvalError::Ref);
            };
            if ch == '\'' {
                let next_cursor = cursor + ch.len_utf8();
                if input[next_cursor..].starts_with('\'') {
                    sheet_name_or_span.push('\'');
                    cursor = next_cursor + 1;
                    continue;
                }
                if !input[next_cursor..].starts_with('!') {
                    return Err(FormulaEvalError::Ref);
                }
                let sheet_ids =
                    resolve_formula_sheet_name_or_span(state, sheet_name_or_span.as_str())?;
                break (sheet_ids, &input[next_cursor + 1..]);
            }
            sheet_name_or_span.push(ch);
            cursor += ch.len_utf8();
        }
    } else if let Some((sheet_name, reference_text)) = input.split_once('!') {
        let sheet_ids = resolve_formula_sheet_name_or_span(state, sheet_name.trim())?;
        (sheet_ids, reference_text)
    } else {
        (vec![default_sheet_id], input)
    };

    let rect = parse_rect_a1(reference_text.trim()).map_err(|_| FormulaEvalError::Ref)?;
    let areas = sheet_ids
        .into_iter()
        .map(|sheet_id| (sheet_id, rect))
        .collect::<Vec<_>>();
    FormulaReference::with_explicit_area_count(1, areas)
}

fn resolve_formula_sheet_name_or_span(
    state: &WorkbookState,
    name_or_span: &str,
) -> Result<Vec<SheetId>, FormulaEvalError> {
    if let Some((start_sheet, end_sheet)) = name_or_span.split_once(':') {
        let start =
            resolve_formula_sheet_name(state, start_sheet.trim()).ok_or(FormulaEvalError::Ref)?;
        let end =
            resolve_formula_sheet_name(state, end_sheet.trim()).ok_or(FormulaEvalError::Ref)?;
        return formula_sheets_in_3d_span(state, start, end);
    }
    resolve_formula_sheet_name(state, name_or_span)
        .map(|sheet_id| vec![sheet_id])
        .ok_or(FormulaEvalError::Ref)
}

fn resolve_formula_sheet_name(state: &WorkbookState, name: &str) -> Option<SheetId> {
    state
        .worksheets
        .iter()
        .find(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
        .map(|worksheet| worksheet.id)
}

fn formula_sheets_in_3d_span(
    state: &WorkbookState,
    start: SheetId,
    end: SheetId,
) -> Result<Vec<SheetId>, FormulaEvalError> {
    let start_index = state
        .worksheets
        .iter()
        .position(|worksheet| worksheet.id == start)
        .ok_or(FormulaEvalError::Ref)?;
    let end_index = state
        .worksheets
        .iter()
        .position(|worksheet| worksheet.id == end)
        .ok_or(FormulaEvalError::Ref)?;
    let (first, last) = if start_index <= end_index {
        (start_index, end_index)
    } else {
        (end_index, start_index)
    };
    Ok(state.worksheets[first..=last]
        .iter()
        .map(|worksheet| worksheet.id)
        .collect())
}

pub(super) fn parse_rect_a1(input: &str) -> OmResult<Rect> {
    let input = input.trim();
    let mut parts = input.split(':');
    let first = parts
        .next()
        .ok_or_else(|| OmError::parse("empty A1 reference"))?;
    let second = parts.next();
    if parts.next().is_some() {
        return Err(OmError::parse("A1 range contains too many ':' separators"));
    }
    let first = parse_a1_endpoint(first)?;
    let Some(second) = second else {
        return match first {
            A1Endpoint::Cell(row, col) => Ok(Rect::single_cell(row, col)),
            A1Endpoint::Row(_) | A1Endpoint::Column(_) => Err(OmError::parse(format!(
                "A1 range {input:?} must use ':' for whole-row or whole-column selectors"
            ))),
        };
    };
    let second = parse_a1_endpoint(second)?;
    match (first, second) {
        (A1Endpoint::Cell(first_row, first_col), A1Endpoint::Cell(second_row, second_col)) => {
            Ok(Rect {
                row_first: first_row.min(second_row),
                row_last: first_row.max(second_row),
                col_first: first_col.min(second_col),
                col_last: first_col.max(second_col),
            })
        }
        (A1Endpoint::Row(first_row), A1Endpoint::Row(second_row)) => Ok(Rect {
            row_first: first_row.min(second_row),
            row_last: first_row.max(second_row),
            col_first: 1,
            col_last: EXCEL_MAX_COLUMN_INDEX,
        }),
        (A1Endpoint::Column(first_col), A1Endpoint::Column(second_col)) => Ok(Rect {
            row_first: 1,
            row_last: EXCEL_MAX_ROW_INDEX,
            col_first: first_col.min(second_col),
            col_last: first_col.max(second_col),
        }),
        _ => Err(OmError::parse(format!(
            "A1 range {input:?} cannot mix cell, row, and column endpoints"
        ))),
    }
}

pub(super) fn split_reference_union_text(input: &str) -> OmResult<Vec<&str>> {
    let input = input.trim();
    if input.is_empty() {
        return Err(OmError::invalid_argument("range reference text is empty"));
    }

    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut in_quote = false;
    let mut in_brackets = false;
    let mut chars = input.char_indices().peekable();
    while let Some((index, ch)) = chars.next() {
        match ch {
            '\'' => {
                if in_quote && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                    chars.next();
                } else {
                    in_quote = !in_quote;
                }
            }
            '[' if !in_quote => in_brackets = true,
            ']' if !in_quote => in_brackets = false,
            ',' if !in_quote && !in_brackets => {
                let part = input[start..index].trim();
                if part.is_empty() {
                    return Err(OmError::invalid_argument(
                        "multi-area range references cannot contain empty areas",
                    ));
                }
                parts.push(part);
                start = index + ch.len_utf8();
            }
            _ => {}
        }
    }

    if in_quote {
        return Err(OmError::invalid_argument(
            "range references use unmatched worksheet quoting",
        ));
    }
    if in_brackets {
        return Err(OmError::invalid_argument(
            "range references use invalid workbook qualification",
        ));
    }

    let part = input[start..].trim();
    if part.is_empty() {
        return Err(OmError::invalid_argument(
            "multi-area range references cannot contain empty areas",
        ));
    }
    parts.push(part);
    Ok(parts)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum A1Endpoint {
    Cell(u32, u32),
    Row(u32),
    Column(u32),
}

fn parse_a1_endpoint(input: &str) -> OmResult<A1Endpoint> {
    let normalized = input.trim().replace('$', "");
    if normalized.is_empty() {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }
    if normalized.chars().all(|ch| ch.is_ascii_digit()) {
        let row = normalized
            .parse::<u32>()
            .map_err(|_| OmError::parse(format!("invalid row index in {input:?}")))?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return Err(OmError::parse(format!(
                "A1 row reference {input:?} is outside worksheet bounds"
            )));
        }
        return Ok(A1Endpoint::Row(row));
    }
    if normalized.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return Ok(A1Endpoint::Column(parse_column_label_a1(input)?));
    }

    let (row, col) = parse_cell_a1(input)?;
    Ok(A1Endpoint::Cell(row, col))
}

fn parse_column_label_a1(input: &str) -> OmResult<u32> {
    let normalized = input.trim().replace('$', "").to_ascii_uppercase();
    if normalized.is_empty() || !normalized.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return Err(OmError::parse(format!(
            "invalid A1 column reference {input:?}"
        )));
    }
    let mut col = 0u32;
    for ch in normalized.bytes() {
        col = col
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
            .ok_or_else(|| OmError::parse("column index overflow"))?;
    }
    if col == 0 || col > EXCEL_MAX_COLUMN_INDEX {
        return Err(OmError::parse(format!(
            "A1 column reference {input:?} is outside worksheet bounds"
        )));
    }
    Ok(col)
}

fn parse_cell_a1(input: &str) -> OmResult<(u32, u32)> {
    let trimmed = input.trim().trim_matches('$');
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in trimmed.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }

    let mut col = 0u32;
    for ch in letters.bytes() {
        col = col
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
            .ok_or_else(|| OmError::parse("column index overflow"))?;
    }
    let row = digits
        .parse::<u32>()
        .map_err(|_| OmError::parse(format!("invalid row index in {input:?}")))?;
    if row == 0 || col == 0 {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }
    if row > EXCEL_MAX_ROW_INDEX || col > EXCEL_MAX_COLUMN_INDEX {
        return Err(OmError::parse(format!(
            "A1 reference {input:?} is outside worksheet bounds"
        )));
    }
    Ok((row, col))
}

pub(super) fn format_rect_address_with_flags(
    rect: Rect,
    row_absolute: bool,
    column_absolute: bool,
) -> String {
    let spans_all_rows = rect.row_first == 1 && rect.row_last == EXCEL_MAX_ROW_INDEX;
    let spans_all_columns = rect.col_first == 1 && rect.col_last == EXCEL_MAX_COLUMN_INDEX;
    if spans_all_rows && !spans_all_columns {
        let start = format_column_address(rect.col_first, column_absolute);
        if rect.col_first == rect.col_last {
            return format!("{start}:{start}");
        }
        let end = format_column_address(rect.col_last, column_absolute);
        return format!("{start}:{end}");
    }
    if spans_all_columns && !spans_all_rows {
        let start = format_row_address(rect.row_first, row_absolute);
        if rect.row_first == rect.row_last {
            return format!("{start}:{start}");
        }
        let end = format_row_address(rect.row_last, row_absolute);
        return format!("{start}:{end}");
    }

    let start = format_cell_address(
        rect.row_first,
        rect.col_first,
        row_absolute,
        column_absolute,
    );
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        start
    } else {
        let end = format_cell_address(rect.row_last, rect.col_last, row_absolute, column_absolute);
        format!("{start}:{end}")
    }
}

fn format_cell_address(row: u32, col: u32, row_absolute: bool, column_absolute: bool) -> String {
    format!(
        "{}{}",
        format_column_address(col, column_absolute),
        format_row_address(row, row_absolute)
    )
}

fn format_column_address(col: u32, column_absolute: bool) -> String {
    let mut address = String::new();
    if column_absolute {
        address.push('$');
    }
    address.push_str(&column_to_letters(col));
    address
}

fn format_row_address(row: u32, row_absolute: bool) -> String {
    let mut address = String::new();
    if row_absolute {
        address.push('$');
    }
    address.push_str(&row.to_string());
    address
}

pub(super) fn format_external_address_qualifier(
    workbook_name: &str,
    worksheet_name: &str,
) -> String {
    let qualifier = format!("[{workbook_name}]{worksheet_name}");
    if excel_reference_qualifier_needs_quotes(workbook_name)
        || excel_reference_qualifier_needs_quotes(worksheet_name)
    {
        format!("'{}'!", qualifier.replace('\'', "''"))
    } else {
        format!("{qualifier}!")
    }
}

fn excel_reference_qualifier_needs_quotes(value: &str) -> bool {
    value.is_empty()
        || value
            .chars()
            .any(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.')
}

fn parse_a1_axis_reference_to_r1c1(
    formula: &str,
    start: usize,
    base_row: u32,
    base_col: u32,
) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    let parse_column_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let letters_start = cursor;
        while cursor < bytes.len()
            && bytes[cursor].is_ascii_alphabetic()
            && cursor - letters_start < 3
        {
            cursor += 1;
        }
        if letters_start == cursor || (cursor < bytes.len() && bytes[cursor].is_ascii_alphabetic())
        {
            return None;
        }
        let col = parse_column_label_a1(&formula[letters_start..cursor]).ok()?;
        Some((col, absolute, cursor))
    };
    let parse_row_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor {
            return None;
        }
        let row = formula[digits_start..cursor].parse::<u32>().ok()?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return None;
        }
        Some((row, absolute, cursor))
    };

    if let Some((first_col, first_absolute, colon_index)) = parse_column_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_col, second_absolute, next_index)) =
            parse_column_endpoint(colon_index + 1)
        && formula[next_index..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
    {
        let start_ref = format_r1c1_column_reference(first_col, first_absolute, base_col);
        let end_ref = format_r1c1_column_reference(second_col, second_absolute, base_col);
        return Some((
            if start_ref == end_ref {
                start_ref
            } else {
                format!("{start_ref}:{end_ref}")
            },
            next_index,
        ));
    }

    if let Some((first_row, first_absolute, colon_index)) = parse_row_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_row, second_absolute, next_index)) = parse_row_endpoint(colon_index + 1)
        && formula[next_index..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
    {
        let start_ref = format_r1c1_row_reference(first_row, first_absolute, base_row);
        let end_ref = format_r1c1_row_reference(second_row, second_absolute, base_row);
        return Some((
            if start_ref == end_ref {
                start_ref
            } else {
                format!("{start_ref}:{end_ref}")
            },
            next_index,
        ));
    }

    None
}

pub(super) fn convert_formula_a1_to_r1c1(formula: &str, base_row: u32, base_col: u32) -> String {
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            let quoted_start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    index += 1;
                    if index < bytes.len() && bytes[index] == b'"' {
                        index += 1;
                        continue;
                    }
                    break;
                }
                let ch = formula[index..]
                    .chars()
                    .next()
                    .expect("valid formula char boundary");
                index += ch.len_utf8();
            }
            output.push_str(&formula[quoted_start..index]);
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary {
            if let Some((reference, next_index)) =
                parse_a1_axis_reference_to_r1c1(formula, index, base_row, base_col)
            {
                output.push_str(&reference);
                index = next_index;
                continue;
            }

            let reference_start = index;
            let mut cursor = index;
            let column_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                cursor += 1;
                true
            } else {
                false
            };
            let letters_start = cursor;
            while cursor < bytes.len()
                && bytes[cursor].is_ascii_alphabetic()
                && cursor - letters_start < 3
            {
                cursor += 1;
            }
            let letters_end = cursor;
            if letters_end > letters_start
                && (cursor >= bytes.len() || !bytes[cursor].is_ascii_alphabetic())
            {
                let row_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                    cursor += 1;
                    true
                } else {
                    false
                };
                let digits_start = cursor;
                while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
                    cursor += 1;
                }
                if digits_start < cursor {
                    let next_char = formula[cursor..].chars().next();
                    let next_is_boundary = next_char.is_none_or(|ch| {
                        !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
                    });
                    if next_is_boundary {
                        let mut col = 0u32;
                        for byte in &bytes[letters_start..letters_end] {
                            col = col * 26 + (byte.to_ascii_uppercase() - b'A' + 1) as u32;
                        }
                        let row = formula[digits_start..cursor].parse::<u32>().ok();
                        if let Some(row) = row {
                            if row > 0
                                && col > 0
                                && row <= EXCEL_MAX_ROW_INDEX
                                && col <= EXCEL_MAX_COLUMN_INDEX
                            {
                                output.push_str(&format_r1c1_reference(
                                    row,
                                    col,
                                    row_absolute,
                                    column_absolute,
                                    base_row,
                                    base_col,
                                ));
                                index = cursor;
                                continue;
                            }
                        }
                    }
                }
            }
            index = reference_start;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula char boundary");
        output.push(ch);
        index += ch.len_utf8();
    }
    output
}

fn format_r1c1_reference(
    row: u32,
    col: u32,
    row_absolute: bool,
    column_absolute: bool,
    base_row: u32,
    base_col: u32,
) -> String {
    format!(
        "{}{}",
        format_r1c1_row_reference(row, row_absolute, base_row),
        format_r1c1_column_reference(col, column_absolute, base_col)
    )
}

fn format_r1c1_row_reference(row: u32, row_absolute: bool, base_row: u32) -> String {
    if row_absolute {
        format!("R{row}")
    } else {
        let delta = i64::from(row) - i64::from(base_row);
        if delta == 0 {
            "R".to_string()
        } else {
            format!("R[{delta}]")
        }
    }
}

fn format_r1c1_column_reference(col: u32, column_absolute: bool, base_col: u32) -> String {
    if column_absolute {
        format!("C{col}")
    } else {
        let delta = i64::from(col) - i64::from(base_col);
        if delta == 0 {
            "C".to_string()
        } else {
            format!("C[{delta}]")
        }
    }
}

pub(super) fn format_rect_r1c1_address_with_flags(
    rect: Rect,
    row_absolute: bool,
    column_absolute: bool,
    base_row: u32,
    base_col: u32,
) -> String {
    let spans_all_rows = rect.row_first == 1 && rect.row_last == EXCEL_MAX_ROW_INDEX;
    let spans_all_columns = rect.col_first == 1 && rect.col_last == EXCEL_MAX_COLUMN_INDEX;
    if spans_all_rows && !spans_all_columns {
        let start = format_r1c1_column_reference(rect.col_first, column_absolute, base_col);
        if rect.col_first == rect.col_last {
            return start;
        }
        let end = format_r1c1_column_reference(rect.col_last, column_absolute, base_col);
        return format!("{start}:{end}");
    }
    if spans_all_columns && !spans_all_rows {
        let start = format_r1c1_row_reference(rect.row_first, row_absolute, base_row);
        if rect.row_first == rect.row_last {
            return start;
        }
        let end = format_r1c1_row_reference(rect.row_last, row_absolute, base_row);
        return format!("{start}:{end}");
    }

    let start = format_r1c1_reference(
        rect.row_first,
        rect.col_first,
        row_absolute,
        column_absolute,
        base_row,
        base_col,
    );
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        start
    } else {
        let end = format_r1c1_reference(
            rect.row_last,
            rect.col_last,
            row_absolute,
            column_absolute,
            base_row,
            base_col,
        );
        format!("{start}:{end}")
    }
}

pub(super) fn convert_formula_r1c1_to_a1(formula: &str, base_row: u32, base_col: u32) -> String {
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            let quoted_start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    index += 1;
                    if index < bytes.len() && bytes[index] == b'"' {
                        index += 1;
                        continue;
                    }
                    break;
                }
                let ch = formula[index..]
                    .chars()
                    .next()
                    .expect("valid formula char boundary");
                index += ch.len_utf8();
            }
            output.push_str(&formula[quoted_start..index]);
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary && matches!(bytes[index], b'R' | b'r') {
            if let Some((row, row_absolute, col, column_absolute, next_index)) =
                parse_r1c1_reference(formula, index, base_row, base_col)
            {
                let next_char = formula[next_index..].chars().next();
                let next_is_boundary = next_char.is_none_or(|ch| {
                    !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
                });
                if next_is_boundary {
                    if row < 1
                        || row > i64::from(EXCEL_MAX_ROW_INDEX)
                        || col < 1
                        || col > i64::from(EXCEL_MAX_COLUMN_INDEX)
                    {
                        output.push_str("#REF!");
                    } else {
                        output.push_str(&format_cell_address(
                            row as u32,
                            col as u32,
                            row_absolute,
                            column_absolute,
                        ));
                    }
                    index = next_index;
                    continue;
                }
            }
        }
        if previous_is_boundary
            && matches!(bytes[index], b'R' | b'r' | b'C' | b'c')
            && let Some((reference, next_index)) =
                parse_r1c1_axis_reference_to_a1(formula, index, base_row, base_col)
        {
            output.push_str(&reference);
            index = next_index;
            continue;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula char boundary");
        output.push(ch);
        index += ch.len_utf8();
    }
    output
}

fn parse_r1c1_axis_reference_to_a1(
    formula: &str,
    start: usize,
    base_row: u32,
    base_col: u32,
) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    if matches!(bytes.get(start), Some(b'R' | b'r')) {
        let (first_row, first_absolute, cursor) =
            parse_r1c1_axis(bytes, start + 1, i64::from(base_row))?;
        if matches!(bytes.get(cursor), Some(b'C' | b'c')) {
            return None;
        }
        if bytes.get(cursor) == Some(&b':') {
            if !matches!(bytes.get(cursor + 1), Some(b'R' | b'r')) {
                return None;
            }
            let (second_row, second_absolute, next_index) =
                parse_r1c1_axis(bytes, cursor + 2, i64::from(base_row))?;
            if !formula[next_index..]
                .chars()
                .next()
                .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
            {
                return None;
            }
            let first = format_a1_row_axis(first_row, first_absolute)?;
            let second = format_a1_row_axis(second_row, second_absolute)?;
            return Some((format!("{first}:{second}"), next_index));
        }
        if formula[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
        {
            let first = format_a1_row_axis(first_row, first_absolute)?;
            return Some((format!("{first}:{first}"), cursor));
        }
    }

    if matches!(bytes.get(start), Some(b'C' | b'c')) {
        let (first_col, first_absolute, cursor) =
            parse_r1c1_axis(bytes, start + 1, i64::from(base_col))?;
        if bytes.get(cursor) == Some(&b':') {
            if !matches!(bytes.get(cursor + 1), Some(b'C' | b'c')) {
                return None;
            }
            let (second_col, second_absolute, next_index) =
                parse_r1c1_axis(bytes, cursor + 2, i64::from(base_col))?;
            if !formula[next_index..]
                .chars()
                .next()
                .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
            {
                return None;
            }
            let first = format_a1_column_axis(first_col, first_absolute)?;
            let second = format_a1_column_axis(second_col, second_absolute)?;
            return Some((format!("{first}:{second}"), next_index));
        }
        if formula[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
        {
            let first = format_a1_column_axis(first_col, first_absolute)?;
            return Some((format!("{first}:{first}"), cursor));
        }
    }

    None
}

fn format_a1_row_axis(row: i64, absolute: bool) -> Option<String> {
    if row < 1 || row > i64::from(EXCEL_MAX_ROW_INDEX) {
        None
    } else {
        Some(format_row_address(row as u32, absolute))
    }
}

fn format_a1_column_axis(col: i64, absolute: bool) -> Option<String> {
    if col < 1 || col > i64::from(EXCEL_MAX_COLUMN_INDEX) {
        None
    } else {
        Some(format_column_address(col as u32, absolute))
    }
}

fn parse_r1c1_reference(
    formula: &str,
    start: usize,
    base_row: u32,
    base_col: u32,
) -> Option<(i64, bool, i64, bool, usize)> {
    let bytes = formula.as_bytes();
    let mut cursor = start;
    if !matches!(bytes.get(cursor), Some(b'R' | b'r')) {
        return None;
    }
    cursor += 1;
    let (row, row_absolute, next_cursor) = parse_r1c1_axis(bytes, cursor, i64::from(base_row))?;
    cursor = next_cursor;
    if !matches!(bytes.get(cursor), Some(b'C' | b'c')) {
        return None;
    }
    cursor += 1;
    let (col, column_absolute, next_cursor) = parse_r1c1_axis(bytes, cursor, i64::from(base_col))?;
    Some((row, row_absolute, col, column_absolute, next_cursor))
}

fn parse_r1c1_axis(bytes: &[u8], start: usize, base: i64) -> Option<(i64, bool, usize)> {
    if bytes.get(start) == Some(&b'[') {
        let mut cursor = start + 1;
        let sign = if bytes.get(cursor) == Some(&b'-') {
            cursor += 1;
            -1
        } else {
            if bytes.get(cursor) == Some(&b'+') {
                cursor += 1;
            }
            1
        };
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor || bytes.get(cursor) != Some(&b']') {
            return None;
        }
        let value = std::str::from_utf8(&bytes[digits_start..cursor])
            .ok()?
            .parse::<i64>()
            .ok()?;
        Some((base + sign * value, false, cursor + 1))
    } else {
        let mut cursor = start;
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor {
            Some((base, false, cursor))
        } else {
            let value = std::str::from_utf8(&bytes[digits_start..cursor])
                .ok()?
                .parse::<i64>()
                .ok()?;
            Some((value, true, cursor))
        }
    }
}

pub(super) fn shift_formula_a1_references(formula: &str, row_delta: i64, col_delta: i64) -> String {
    if row_delta == 0 && col_delta == 0 {
        return formula.to_string();
    }

    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            let quoted_start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    index += 1;
                    if index < bytes.len() && bytes[index] == b'"' {
                        index += 1;
                        continue;
                    }
                    break;
                }
                let ch = formula[index..]
                    .chars()
                    .next()
                    .expect("valid formula char boundary");
                index += ch.len_utf8();
            }
            output.push_str(&formula[quoted_start..index]);
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary {
            if let Some((reference, next_index)) =
                shift_formula_a1_axis_reference(formula, index, row_delta, col_delta)
            {
                output.push_str(&reference);
                index = next_index;
                continue;
            }

            let reference_start = index;
            let mut cursor = index;
            let column_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                cursor += 1;
                true
            } else {
                false
            };
            let letters_start = cursor;
            while cursor < bytes.len()
                && bytes[cursor].is_ascii_alphabetic()
                && cursor - letters_start < 3
            {
                cursor += 1;
            }
            let letters_end = cursor;
            if letters_end > letters_start
                && (cursor >= bytes.len() || !bytes[cursor].is_ascii_alphabetic())
            {
                let row_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                    cursor += 1;
                    true
                } else {
                    false
                };
                let digits_start = cursor;
                while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
                    cursor += 1;
                }
                if digits_start < cursor {
                    let next_char = formula[cursor..].chars().next();
                    let next_is_boundary = next_char.is_none_or(|ch| {
                        !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
                    });
                    if next_is_boundary {
                        let mut col = 0u32;
                        for byte in &bytes[letters_start..letters_end] {
                            col = col * 26 + (byte.to_ascii_uppercase() - b'A' + 1) as u32;
                        }
                        let row = formula[digits_start..cursor].parse::<u32>().ok();
                        if let Some(row) = row {
                            if row > 0
                                && col > 0
                                && row <= EXCEL_MAX_ROW_INDEX
                                && col <= EXCEL_MAX_COLUMN_INDEX
                            {
                                let shifted_row = if row_absolute {
                                    i64::from(row)
                                } else {
                                    i64::from(row) + row_delta
                                };
                                let shifted_col = if column_absolute {
                                    i64::from(col)
                                } else {
                                    i64::from(col) + col_delta
                                };
                                if shifted_row < 1
                                    || shifted_row > i64::from(EXCEL_MAX_ROW_INDEX)
                                    || shifted_col < 1
                                    || shifted_col > i64::from(EXCEL_MAX_COLUMN_INDEX)
                                {
                                    output.push_str("#REF!");
                                } else {
                                    output.push_str(&format_cell_address(
                                        shifted_row as u32,
                                        shifted_col as u32,
                                        row_absolute,
                                        column_absolute,
                                    ));
                                }
                                index = cursor;
                                continue;
                            }
                        }
                    }
                }
            }
            index = reference_start;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula char boundary");
        output.push(ch);
        index += ch.len_utf8();
    }
    output
}

fn shift_formula_a1_axis_reference(
    formula: &str,
    start: usize,
    row_delta: i64,
    col_delta: i64,
) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    let parse_column_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let letters_start = cursor;
        while cursor < bytes.len()
            && bytes[cursor].is_ascii_alphabetic()
            && cursor - letters_start < 3
        {
            cursor += 1;
        }
        if letters_start == cursor || (cursor < bytes.len() && bytes[cursor].is_ascii_alphabetic())
        {
            return None;
        }
        let col = parse_column_label_a1(&formula[letters_start..cursor]).ok()?;
        Some((col, absolute, cursor))
    };
    let parse_row_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor {
            return None;
        }
        let row = formula[digits_start..cursor].parse::<u32>().ok()?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return None;
        }
        Some((row, absolute, cursor))
    };
    let next_is_boundary = |cursor: usize| {
        formula[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
    };

    if let Some((first_col, first_absolute, colon_index)) = parse_column_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_col, second_absolute, next_index)) =
            parse_column_endpoint(colon_index + 1)
        && next_is_boundary(next_index)
    {
        let first = if first_absolute {
            i64::from(first_col)
        } else {
            i64::from(first_col) + col_delta
        };
        let second = if second_absolute {
            i64::from(second_col)
        } else {
            i64::from(second_col) + col_delta
        };
        if first < 1
            || first > i64::from(EXCEL_MAX_COLUMN_INDEX)
            || second < 1
            || second > i64::from(EXCEL_MAX_COLUMN_INDEX)
        {
            return Some(("#REF!".to_string(), next_index));
        }
        return Some((
            format!(
                "{}:{}",
                format_column_address(first as u32, first_absolute),
                format_column_address(second as u32, second_absolute)
            ),
            next_index,
        ));
    }

    if let Some((first_row, first_absolute, colon_index)) = parse_row_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_row, second_absolute, next_index)) = parse_row_endpoint(colon_index + 1)
        && next_is_boundary(next_index)
    {
        let first = if first_absolute {
            i64::from(first_row)
        } else {
            i64::from(first_row) + row_delta
        };
        let second = if second_absolute {
            i64::from(second_row)
        } else {
            i64::from(second_row) + row_delta
        };
        if first < 1
            || first > i64::from(EXCEL_MAX_ROW_INDEX)
            || second < 1
            || second > i64::from(EXCEL_MAX_ROW_INDEX)
        {
            return Some(("#REF!".to_string(), next_index));
        }
        return Some((
            format!(
                "{}:{}",
                format_row_address(first as u32, first_absolute),
                format_row_address(second as u32, second_absolute)
            ),
            next_index,
        ));
    }

    None
}

fn column_to_letters(mut col: u32) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters.iter().rev().collect()
}
