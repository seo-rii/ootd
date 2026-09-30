//! The scalar worksheet function table and its numeric evaluation.

use super::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormulaScalarFunction {
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
    pub(super) fn from_name(name: &str) -> Option<Self> {
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

    pub(super) fn evaluate(
        self,
        args: &[f64],
        context: &CalcContext,
    ) -> Result<f64, FormulaEvalError> {
        let date_system = context.date_system();
        let serial_weekday_monday0 = |serial: i64| {
            let serial = date_system.serial_1900(serial);
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
            let (year, month, day) = date_system.ymd(serial as f64)?;
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
            date_system
                .serial_from_args((year + 1) as f64, 1.0, 1.0)
                .map(|value| value as i64)
        };
        let serial_to_year_start = |year: i64| -> Result<i64, FormulaEvalError> {
            date_system
                .serial_from_args(year as f64, 1.0, 1.0)
                .map(|value| value as i64)
        };
        let days360 =
            |start_serial: i64, end_serial: i64, european: bool| -> Result<i64, FormulaEvalError> {
                let (start_serial, end_serial, sign) = if start_serial > end_serial {
                    (end_serial, start_serial, -1)
                } else {
                    (start_serial, end_serial, 1)
                };
                let (start_year, start_month, start_day) = date_system.ymd(start_serial as f64)?;
                let (mut end_year, mut end_month, mut end_day) =
                    date_system.ymd(end_serial as f64)?;
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
                let (start_year, _, _) = date_system.ymd(start_serial as f64)?;
                let (end_year, _, _) = date_system.ymd(end_serial as f64)?;
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
            date_system
                .ymd(serial as f64)
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
                    date_system.edate(next_coupon as f64, -(months_per_coupon as f64))? as i64;
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
                            date_system.edate(coupon_date as f64, months_per_coupon as f64)? as i64;
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
                date_system.edate(first_coupon as f64, -(months_per_coupon as f64))? as i64;
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
                date_system.edate(first_coupon as f64, months_per_coupon as f64)? as i64;
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
                coupon_date =
                    date_system.edate(coupon_date as f64, months_per_coupon as f64)? as i64;
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
                date_system.edate(last_interest as f64, months_per_coupon as f64)? as i64;
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
                        date_system.edate(next_coupon as f64, months_per_coupon as f64)? as i64;
                    guard += 1;
                    if guard > 10000 {
                        return Err(FormulaEvalError::Num);
                    }
                }
                loop {
                    let previous_coupon =
                        date_system.edate(next_coupon as f64, -(months_per_coupon as f64))? as i64;
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
                        date_system.edate(next_coupon as f64, -(months_per_coupon as f64))? as i64;
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
                        date_system.edate(next_coupon as f64, months_per_coupon as f64)? as i64;
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
                date_system.serial_from_args(*year, *month, *day)
            }
            FormulaScalarFunction::Day => {
                let [serial] = args else {
                    return Err(FormulaEvalError::Value);
                };
                let (_, _, day) = date_system.ymd(*serial)?;
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
                date_system.edate(*serial, *months)
            }
            FormulaScalarFunction::EOMonth => {
                let [serial, months] = args else {
                    return Err(FormulaEvalError::Value);
                };
                date_system.eomonth(*serial, *months)
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
                let (_, month, _) = date_system.ymd(*serial)?;
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
                context.now_serial()
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
                Ok(context.rand())
            }
            FormulaScalarFunction::RandBetween => {
                let [bottom, top] = args else {
                    return Err(FormulaEvalError::Value);
                };
                context.rand_between(*bottom, *top)
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
                context.now_serial().map(f64::floor)
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
                let (year, _, _) = date_system.ymd(serial as f64)?;
                let jan1_serial = date_system.serial_from_args(year as f64, 1.0, 1.0)? as i64;
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
                let (year, _, _) = date_system.ymd(*serial)?;
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
