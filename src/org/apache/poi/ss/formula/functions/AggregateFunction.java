/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.ss.formula.functions;

import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.ValueEval;

/**
 * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
 */
public abstract class AggregateFunction extends MultiOperandNumericFunction {

	private static final class LargeSmall extends Fixed2ArgFunction {
		private final boolean _isLarge;
		protected LargeSmall(boolean isLarge) {
			_isLarge = isLarge;
		}

		public ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0,
				ValueEval arg1) {
			double dn;
			try {
				ValueEval ve1 = OperandResolver.getSingleValue(arg1, srcRowIndex, srcColumnIndex);
				dn = OperandResolver.coerceValueToDouble(ve1);
			} catch (EvaluationException e1) {
				// all errors in the second arg translate to #VALUE!
				return ErrorEval.VALUE_INVALID;
			}
			// weird Excel behaviour on second arg
			if (dn < 1.0) {
				// values between 0.0 and 1.0 result in #NUM!
				return ErrorEval.NUM_ERROR;
			}
			// all other values are rounded up to the next integer
			int k = (int) Math.ceil(dn);

			double result;
			try {
				double[] ds = ValueCollector.collectValues(arg0);
				if (k > ds.length) {
					return ErrorEval.NUM_ERROR;
				}
				result = _isLarge ? StatsLib.kthLargest(ds, k) : StatsLib.kthSmallest(ds, k);
				NumericFunction.checkValue(result);
			} catch (EvaluationException e) {
				return e.getErrorEval();
			}

			return new NumberEval(result);
		}
	}
	private static final class Percentile extends Fixed2ArgFunction {
				
			protected Percentile() {
			}
	
			public ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0,
					ValueEval arg1) {
				double dn;
				try {
					ValueEval ve1 = OperandResolver.getSingleValue(arg1, srcRowIndex, srcColumnIndex);
					dn = OperandResolver.coerceValueToDouble(ve1);
				} catch (EvaluationException e1) {
					// all errors in the second arg translate to #VALUE!
					return ErrorEval.VALUE_INVALID;
				}
				if (dn < 0 || dn > 1) { // has to be percentage
				return ErrorEval.NUM_ERROR;
			}
	
				double result;
				try {
					double[] ds = ValueCollector.collectValues(arg0);
					int N = ds.length;
					double n = (N - 1) * dn + 1;
					if (n == 1d) {
						result = StatsLib.kthSmallest(ds, 1);
					} else if (n == N) {
						result = StatsLib.kthLargest(ds, 1);
					} else {
						int k = (int) n;
						double d = n - k;
						result = StatsLib.kthSmallest(ds, k) + d
								* (StatsLib.kthSmallest(ds, k + 1) - StatsLib.kthSmallest(ds, k));
					}
	
					NumericFunction.checkValue(result);
				} catch (EvaluationException e) {
					return e.getErrorEval();
				}
	
				return new NumberEval(result);
			}
	}
			
	static final class ValueCollector extends MultiOperandNumericFunction {
		private static final ValueCollector instance = new ValueCollector();
		public ValueCollector() {
			super(false, false);
		}
		public static double[] collectValues(ValueEval...operands) throws EvaluationException {
			return instance.getNumberArray(operands);
		}
		protected double evaluate(double[] values) {
			throw new IllegalStateException("should not be called");
		}
	}
	
    protected AggregateFunction() {
        super(false, false);
    }

    /**
     * Create an instance to use in the {@link Subtotal} function.
     *
     * <p>
     *     If there are other subtotals within argument refs (or nested subtotals),
     *     these nested subtotals are ignored to avoid double counting.
     * </p>
     *
     * @param   func  the function to wrap
     * @return  wrapped instance. The actual math is delegated to the argument function.
     */
    /*package*/ static Function subtotalInstance(Function func) {
        final AggregateFunction arg = (AggregateFunction)func;
        return new AggregateFunction() {
            @Override
            protected double evaluate(double[] values) throws EvaluationException {
                return arg.evaluate(values);
            }

            /**
             *  ignore nested subtotals.
             */
            @Override
            public boolean isSubtotalCounted(){
                return false;
            }

        };
    }

    public static final Function AVEDEV = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return StatsLib.avedev(values);
		}
	};
	public static final Function AVERAGE = new AggregateFunction() {
		protected double evaluate(double[] values) throws EvaluationException {
			if (values.length < 1) {
				throw new EvaluationException(ErrorEval.DIV_ZERO);
			}
			return MathX.average(values);
		}
	};
	public static final Function DEVSQ = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return StatsLib.devsq(values);
		}
	};
	public static final Function LARGE = new LargeSmall(true);
	public static final Function MAX = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return values.length > 0 ? MathX.max(values) : 0;
		}
	};
	public static final Function MEDIAN = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return StatsLib.median(values);
		}
	};
	public static final Function MIN = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return values.length > 0 ? MathX.min(values) : 0;
		}
	};
	public static final Function PERCENTILE = new Percentile();

	public static final Function PRODUCT = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return MathX.product(values);
		}
	};
	public static final Function SMALL = new LargeSmall(false);
	public static final Function STDEV = new AggregateFunction() {
		protected double evaluate(double[] values) throws EvaluationException {
			if (values.length < 1) {
				throw new EvaluationException(ErrorEval.DIV_ZERO);
			}
			return StatsLib.stdev(values);
		}
	};
	public static final Function SUM = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return MathX.sum(values);
		}
	};
	public static final Function SUMSQ = new AggregateFunction() {
		protected double evaluate(double[] values) {
			return MathX.sumsq(values);
		}
	};
    public static final Function VAR = new AggregateFunction() {
        protected double evaluate(double[] values) throws EvaluationException {
            if (values.length < 1) {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return StatsLib.var(values);
        }
    };
    public static final Function VARP = new AggregateFunction() {
        protected double evaluate(double[] values) throws EvaluationException {
            if (values.length < 1) {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return StatsLib.varp(values);
        }
    };
}
